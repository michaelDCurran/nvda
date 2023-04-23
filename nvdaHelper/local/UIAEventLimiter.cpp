#include <iostream>
#include <windows.h>
#include <uiautomation.h>
#include <variant>
#include <map>
#include <vector>
#include <functional>
#include <mutex>
#include <atomic>
#include <atlcomcli.h>
#include <comutil.h>

template<typename t> struct debug_print_type;

void log(const char* message) {
	std::cerr<<message<<std::endl;
}

std::vector<int> SafeArrayToVector(SAFEARRAY* pSafeArray) {
	std::vector<int> vec;
	int* data;
	HRESULT hr = SafeArrayAccessData(pSafeArray, (void**)&data);
	if (SUCCEEDED(hr)) {
		LONG lowerBound, upperBound;
		SafeArrayGetLBound(pSafeArray, 1, &lowerBound);
		SafeArrayGetUBound(pSafeArray, 1, &upperBound);
		vec.assign(data, data + (upperBound - lowerBound + 1));
		SafeArrayUnaccessData(pSafeArray);
	}
	return vec;
}

std::vector<int> getRuntimeIDFromElement(IUIAutomationElement* pElement) {
	SAFEARRAY* runtimeIdArray;
	HRESULT hr = pElement->GetRuntimeId(&runtimeIdArray);
	if (FAILED(hr)) {
		return {};
	}
	std::vector<int> runtimeID = SafeArrayToVector(runtimeIdArray);
	SafeArrayDestroy(runtimeIdArray);
	return runtimeID;
}

class EventRecord_t {
	public:
	CComPtr<IUIAutomationElement> element;
	EventRecord_t(IUIAutomationElement* pElement): element(pElement) {
	}
};

class FocusChangedEventRecord_t: public EventRecord_t {
	public:
	FocusChangedEventRecord_t(IUIAutomationElement* pElement): EventRecord_t(pElement) {
	}
	HRESULT emit(IUIAutomationFocusChangedEventHandler* handler) const {
		return handler->HandleFocusChangedEvent(element);
	}
};

class CoalesceableEventRecord_t: public EventRecord_t {
	public:
	std::vector<int> runtimeID;
	unsigned int coalesceCount;
	CoalesceableEventRecord_t(IUIAutomationElement* pElement): EventRecord_t(pElement) {
		runtimeID = getRuntimeIDFromElement(element);
	}
	bool isEqual(const CoalesceableEventRecord_t& other) const {
		return runtimeID == other.runtimeID;
	}
};

class AutomationEventRecord_t: public CoalesceableEventRecord_t {
	public:
	EVENTID eventID;
	AutomationEventRecord_t(IUIAutomationElement* pElement, EVENTID eventID): CoalesceableEventRecord_t(pElement), eventID(eventID) {
	}
	HRESULT emit(IUIAutomationEventHandler* handler) const {
		return handler->HandleAutomationEvent(element, eventID);
	}
	bool isEqual(const AutomationEventRecord_t& other) const {
		return eventID == other.eventID && static_cast<CoalesceableEventRecord_t>(*this).isEqual(other);
	}
};

class PropertyChangedEventRecord_t: public CoalesceableEventRecord_t {
	public:
	PROPERTYID propertyID;
	CComVariant value;
	PropertyChangedEventRecord_t(IUIAutomationElement* pElement, PROPERTYID propertyID, VARIANT value): CoalesceableEventRecord_t(pElement), propertyID(propertyID), value(value) { 
	}
	HRESULT emit(IUIAutomationPropertyChangedEventHandler* handler) const {
		return handler->HandlePropertyChangedEvent(element, propertyID, value);
	}
	bool isEqual(const PropertyChangedEventRecord_t& other) const {
		return propertyID == other.propertyID && static_cast<CoalesceableEventRecord_t>(*this).isEqual(other);
	}
};

using AnyEventRecord_t = std::variant<FocusChangedEventRecord_t, AutomationEventRecord_t, PropertyChangedEventRecord_t>;

class RateLimitedEventHandler : public IUIAutomationEventHandler, public IUIAutomationFocusChangedEventHandler, public IUIAutomationPropertyChangedEventHandler {
private:
	unsigned long m_refCount;
	CComQIPtr<IUIAutomationEventHandler> m_pExistingAutomationEventHandler;
	CComQIPtr<IUIAutomationFocusChangedEventHandler> m_pExistingFocusChangedEventHandler;
	CComQIPtr<IUIAutomationPropertyChangedEventHandler> m_pExistingPropertyChangedEventHandler;
	void(*m_onFirstEvent)(void);
	std::mutex mtx;
	std::list<AnyEventRecord_t> m_eventRecords;

	HRESULT addEvent(AnyEventRecord_t&& recordVar) {
		log("RateLimitedUIAEventHandler::addEvent called");
		bool needsCallback = false;
		{ std::lock_guard lock(mtx);
			if(m_eventRecords.empty()) {
				log("RateLimitedUIAEventHandler::addEvent: First event, needs callback.");
				needsCallback = true;
			}
			unsigned int coalesceCount = 0;
			std::visit([this,&coalesceCount](const auto& record) {
				using RecordType = std::decay_t<decltype(record)>;
				if constexpr(std::is_base_of_v<CoalesceableEventRecord_t, RecordType>) {
					log("RateLimitedUIAEventHandler::addEvent: Is a coalesceable event");
					auto existingIter = std::find_if(m_eventRecords.begin(), m_eventRecords.end(), [&](auto& existingRecordVar) {
						auto existingRecordPtr = std::get_if<RecordType>(&existingRecordVar);
						if(existingRecordPtr) {
							log("RateLimitedUIAEventHandler::addEvent: found an existing event"); 
							coalesceCount += existingRecordPtr->coalesceCount;
							return true;
						}
						return false;
					});
					if(existingIter != m_eventRecords.end()) {
						log("RateLimitedUIAEventHandler::addEvent: removing existing event"); 
						m_eventRecords.erase(existingIter);
					}
				}
			}, recordVar);
			log("RateLimitedUIAEventHandler::addEvent: Inserting new event");
			m_eventRecords.push_back(recordVar);
		}
		if(needsCallback) {
			log("RateLimitedUIAEventHandler::addEvent: Firing callback");
			m_onFirstEvent();
		}
		return S_OK;
	}

	~RateLimitedEventHandler() {
		log("RateLimitedUIAEventHandler::~RateLimitedUIAEventHandler called");
	}

public:

	RateLimitedEventHandler(IUnknown* pExistingHandler, void(*onFirstEvent)(void))
		: m_refCount(1), m_pExistingAutomationEventHandler(pExistingHandler), m_pExistingFocusChangedEventHandler(pExistingHandler), m_pExistingPropertyChangedEventHandler(pExistingHandler) {
			log("RateLimitedUIAEventHandler::RateLimitedUIAEventHandler called");
		}

	// IUnknown methods
	ULONG STDMETHODCALLTYPE AddRef() {
		return InterlockedIncrement(&m_refCount);
	}

	ULONG STDMETHODCALLTYPE Release() {
		ULONG refCount = InterlockedDecrement(&m_refCount);
		if (refCount == 0) {
			delete this;
		}
		return refCount;
	}

	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface) {
		if (riid == __uuidof(IUnknown)) {
			*ppInterface = static_cast<IUIAutomationEventHandler*>(this);
			AddRef();
			return S_OK;
		} else if (riid == __uuidof(IUIAutomationEventHandler)) {
			*ppInterface = static_cast<IUIAutomationEventHandler*>(this);
			AddRef();
			return S_OK;
		} else if (riid == __uuidof(IUIAutomationFocusChangedEventHandler)) {
			*ppInterface = static_cast<IUIAutomationFocusChangedEventHandler*>(this);
			AddRef();
			return S_OK;
		} else if (riid == __uuidof(IUIAutomationPropertyChangedEventHandler)) {
			*ppInterface = static_cast<IUIAutomationPropertyChangedEventHandler*>(this);
			AddRef();
			return S_OK;
		}
		*ppInterface = nullptr;
		return E_NOINTERFACE;
	}

	// IUIAutomationEventHandler method
	HRESULT STDMETHODCALLTYPE HandleAutomationEvent(IUIAutomationElement* pElement, EVENTID eventID) {
		log("RateLimitedUIAEventHandler::HandleAutomationEvent called");
		if(!m_pExistingAutomationEventHandler) {
			log("RateLimitedUIAEventHandler::HandleAutomationEvent: No existing evnet handler. Returning");
			return E_NOTIMPL;
		}
		return addEvent(AutomationEventRecord_t(pElement, eventID));
	}

	// IUIAutomationFocusEventHandler method
	HRESULT STDMETHODCALLTYPE HandleFocusChangedEvent(IUIAutomationElement* pElement) {
		log("RateLimitedUIAEventHandler::HandleFocusChangedEvent called");
		if(!m_pExistingFocusChangedEventHandler) {
			log("RateLimitedUIAEventHandler::HandleFocusChangedEvent: No existing focusChangeEventHandler, returning");
			return E_NOTIMPL;
		}
		return addEvent(FocusChangedEventRecord_t(pElement));
	}

	// IUIAutomationPropertyChangedEventHandler method
	HRESULT STDMETHODCALLTYPE HandlePropertyChangedEvent(
	IUIAutomationElement* pElement, PROPERTYID propertyID, VARIANT newValue) {
		log("RateLimitedUIAEventHandler::HandlePropertyChangedEvent called");
		if(!m_pExistingPropertyChangedEventHandler) {
			log("RateLimitedUIAEventHandler::HandlePropertyChangedEvent: no existing handler. Returning");
			return E_NOTIMPL;
		}
		return addEvent(PropertyChangedEventRecord_t(pElement, propertyID, newValue));
	}

	void flush() {
		log("RateLimitedUIAEventHandler::flush called");
		std::list<AnyEventRecord_t> eventRecordsCopy;
		{ std::lock_guard lock(mtx);
			eventRecordsCopy.swap(m_eventRecords);
		}

		// Emit events
		for(auto& recordVar: eventRecordsCopy) {
			std::visit([this](const auto& record) {
				using RecordType = std::decay_t<decltype(record)>;
				if constexpr(std::is_same_v<FocusChangedEventRecord_t, RecordType>) {
					record.emit(this->m_pExistingFocusChangedEventHandler);
				} else if constexpr(std::is_same_v<AutomationEventRecord_t, RecordType>) {
					record.emit(this->m_pExistingAutomationEventHandler);
				} else if constexpr(std::is_same_v<PropertyChangedEventRecord_t, RecordType>) {
					record.emit(this->m_pExistingPropertyChangedEventHandler);
				} else {
					debug_print_type<RecordType>{};
				}
			}, recordVar);
		}
	}

};

HRESULT rateLimitedUIAEventHandler_create(IUnknown* pExistingHandler, void(*onFirstEvent)(void), RateLimitedEventHandler** ppRateLimitedEventHandler) {
	log("rateLimitedUIAEventHandler_create called");
	if (!pExistingHandler || !ppRateLimitedEventHandler) {
		log("rateLimitedUIAEventHandler_create: invalid arguments. Returning");
		return E_INVALIDARG;
	}

	// Create the RateLimitedEventHandler instance
	*ppRateLimitedEventHandler = new RateLimitedEventHandler(pExistingHandler, onFirstEvent);
	if (!(*ppRateLimitedEventHandler)) {
		log("rateLimitedUIAEventHandler_create: Could not create RateLimitedUIAEventHandler. Returning");
		return E_OUTOFMEMORY;
	}
	log("rateLimitedUIAEventHandler_create: done");
	return S_OK;
}

HRESULT rateLimitedUIAEventHandler_flush(RateLimitedEventHandler* pRateLimitedEventHandler) {
	log("rateLimitedUIAEventHandler_flush called");
	if(!pRateLimitedEventHandler) {
		log("rateLimitedUIAEventHandler_flush: invalid argument. Returning");
		return E_INVALIDARG;
	}
	pRateLimitedEventHandler->flush();
	return S_OK;
}

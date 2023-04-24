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
#include <common/log.h>

template<typename t> struct debug_print_type;

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

class RateLimitedEventHandler;

class EventRecord_t {
	public:
	CComPtr<IUIAutomationElement> element;
	bool isCoalesceable;
	std::vector<int> coalescingKey;
	unsigned int coalesceCount = 0;
	EventRecord_t(IUIAutomationElement* pElement, bool isCoAlesceable): element(pElement), isCoalesceable(isCoalesceable) {
		if(isCoalesceable) {
			coalescingKey = getRuntimeIDFromElement(element);
		}
	}
	bool canCoalesce(const EventRecord_t& other) CONST {
		return isCoalesceable && other.isCoalesceable && coalescingKey == other.coalescingKey;
	}
	virtual HRESULT emit(const RateLimitedEventHandler& handler) const = 0; 
	virtual ~EventRecord_t() = default;
};

class FocusChangedEventRecord_t: public EventRecord_t {
	public:
	FocusChangedEventRecord_t(IUIAutomationElement* pElement): EventRecord_t(pElement, false) {
	}
	HRESULT emit(const RateLimitedEventHandler& handler) const override;
};

class AutomationEventRecord_t: public EventRecord_t {
	public:
	EVENTID eventID;
	AutomationEventRecord_t(IUIAutomationElement* pElement, EVENTID eventID): EventRecord_t(pElement, true), eventID(eventID) {
		coalescingKey.push_back(eventID);
	}
	HRESULT emit(CONST RateLimitedEventHandler& handler) const override; 
};

class PropertyChangedEventRecord_t: public EventRecord_t {
	public:
	PROPERTYID propertyID;
	CComVariant value;
	PropertyChangedEventRecord_t(IUIAutomationElement* pElement, PROPERTYID propertyID, VARIANT value): EventRecord_t(pElement, true), propertyID(propertyID), value(value) { 
		coalescingKey.push_back(UIA_AutomationPropertyChangedEventId);
		coalescingKey.push_back(propertyID);
	}
	HRESULT emit(CONST RateLimitedEventHandler& handler) const override;
};

class RateLimitedEventHandler : public IUIAutomationEventHandler, public IUIAutomationFocusChangedEventHandler, public IUIAutomationPropertyChangedEventHandler {
private:
	unsigned long m_refCount;
	CComQIPtr<IUIAutomationEventHandler> m_pExistingAutomationEventHandler;
	CComQIPtr<IUIAutomationFocusChangedEventHandler> m_pExistingFocusChangedEventHandler;
	CComQIPtr<IUIAutomationPropertyChangedEventHandler> m_pExistingPropertyChangedEventHandler;
	std::function<void()> m_onFirstEvent;
	std::mutex mtx;
	std::list<std::unique_ptr<EventRecord_t>> m_eventRecords;

	HRESULT queueEvent(std::unique_ptr<EventRecord_t> record) {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent called");
		if(!record) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: NULL record. Returning");
			return E_INVALIDARG;
		}

		bool needsCallback = false;
		{ std::lock_guard lock(mtx);
			if(m_eventRecords.empty()) {
				LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: First event, needs callback.");
				needsCallback = true;
			}
			if(record->isCoalesceable) {
				LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: Is a coalesceable event");
				auto existingIter = std::find_if(m_eventRecords.begin(), m_eventRecords.end(), [&record](const auto& existingRecord) {
					if(record->canCoalesce(*existingRecord)) {
						LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: found an existing event"); 
						record->coalesceCount += existingRecord->coalesceCount;
						return true;
					}
					return false;
				});
				if(existingIter != m_eventRecords.end()) {
					LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: removing existing event"); 
					m_eventRecords.erase(existingIter);
				}
			}
			LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: Inserting new event");
			record->coalesceCount += 1;
			m_eventRecords.push_back(std::move(record));
		}
		if(needsCallback) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: Firing callback");
			m_onFirstEvent();
			LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: Done firing callback");
		}
		return S_OK;
	}

	~RateLimitedEventHandler() {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::~RateLimitedUIAEventHandler called");
	}

	protected:
	friend EventRecord_t;
	friend AutomationEventRecord_t;
	friend FocusChangedEventRecord_t;
	friend PropertyChangedEventRecord_t;

	HRESULT emitAutomationEvent(const AutomationEventRecord_t& record) const {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::emitAutomationEvent called");
		return m_pExistingAutomationEventHandler->HandleAutomationEvent(record.element, record.eventID);
	}

	HRESULT emitFocusChangedEvent(const FocusChangedEventRecord_t& record) const {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::emitFocusChangedEvent called");
		return m_pExistingFocusChangedEventHandler->HandleFocusChangedEvent(record.element);
	}

	HRESULT emitPropertyChangedEvent(const PropertyChangedEventRecord_t& record) const {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::emitPropertyChangedEvent called");
		return m_pExistingPropertyChangedEventHandler->HandlePropertyChangedEvent(record.element, record.propertyID, record.value);
	}

public:

	RateLimitedEventHandler(IUnknown* pExistingHandler, std::function<void()> onFirstEvent)
		: m_onFirstEvent(onFirstEvent), m_refCount(1), m_pExistingAutomationEventHandler(pExistingHandler), m_pExistingFocusChangedEventHandler(pExistingHandler), m_pExistingPropertyChangedEventHandler(pExistingHandler) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::RateLimitedUIAEventHandler called");
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
		LOG_DEBUG(L"RateLimitedUIAEventHandler::HandleAutomationEvent called");
		if(!m_pExistingAutomationEventHandler) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::HandleAutomationEvent: No existing evnet handler. Returning");
			return E_NOTIMPL;
		}
		auto record = std::make_unique<AutomationEventRecord_t>(pElement, eventID);
		return queueEvent(std::move(record));
	}

	// IUIAutomationFocusEventHandler method
	HRESULT STDMETHODCALLTYPE HandleFocusChangedEvent(IUIAutomationElement* pElement) {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::HandleFocusChangedEvent called");
		if(!m_pExistingFocusChangedEventHandler) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::HandleFocusChangedEvent: No existing focusChangeEventHandler, returning");
			return E_NOTIMPL;
		}
		auto record = std::make_unique<FocusChangedEventRecord_t>(pElement);
		return queueEvent(std::move(record));
	}

	// IUIAutomationPropertyChangedEventHandler method
	HRESULT STDMETHODCALLTYPE HandlePropertyChangedEvent(
	IUIAutomationElement* pElement, PROPERTYID propertyID, VARIANT newValue) {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::HandlePropertyChangedEvent called");
		if(!m_pExistingPropertyChangedEventHandler) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::HandlePropertyChangedEvent: no existing handler. Returning");
			return E_NOTIMPL;
		}
		auto record = std::make_unique<PropertyChangedEventRecord_t>(pElement, propertyID, newValue);
		return queueEvent(std::move(record));
	}

	void flush() {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::flush called");
		std::list<std::unique_ptr<EventRecord_t>> eventRecordsCopy;
		{ std::lock_guard lock(mtx);
			eventRecordsCopy.swap(m_eventRecords);
		}

		// Emit events
		LOG_DEBUG(L"RateLimitedUIAEventHandler::flush: Emitting events...");
		for(const auto& record: eventRecordsCopy) {
			record->emit(*this);
		}
		LOG_DEBUG(L"RateLimitedUIAEventHandler::flush: done emitting events"); 
	}

};

HRESULT FocusChangedEventRecord_t::emit(const RateLimitedEventHandler& handler) const {
	LOG_DEBUG(L"FocusChangedEventRecord_t::emit: Emmiting event");
	return handler.emitFocusChangedEvent(*this);
}

HRESULT AutomationEventRecord_t::emit(CONST RateLimitedEventHandler& handler) const {
	LOG_DEBUG(L"AutomationEventRecord_t::emit: Emitting event of type "<<(this->eventID)<<L", coalesceCount "<<(this->coalesceCount));
	if(this->coalesceCount > 1) Beep(500+(this->coalesceCount)*100, 40);
	return handler.emitAutomationEvent(*this);
}

HRESULT PropertyChangedEventRecord_t::emit(CONST RateLimitedEventHandler& handler) const {
	LOG_DEBUG(L"PropertyChangedEventRecord_t::emit: Emitting property changed event for property  "<<(this->propertyID)<<L", coalesceCount "<<(this->coalesceCount));
	return handler.emitPropertyChangedEvent(*this);
}

HRESULT rateLimitedUIAEventHandler_create(IUnknown* pExistingHandler, void(*onFirstEvent)(void), RateLimitedEventHandler** ppRateLimitedEventHandler) {
	LOG_DEBUG(L"rateLimitedUIAEventHandler_create called");
	if (!pExistingHandler || !ppRateLimitedEventHandler) {
		LOG_DEBUG(L"rateLimitedUIAEventHandler_create: invalid arguments. Returning");
		return E_INVALIDARG;
	}

	// Create the RateLimitedEventHandler instance
	*ppRateLimitedEventHandler = new RateLimitedEventHandler(pExistingHandler, onFirstEvent);
	if (!(*ppRateLimitedEventHandler)) {
		LOG_DEBUG(L"rateLimitedUIAEventHandler_create: Could not create RateLimitedUIAEventHandler. Returning");
		return E_OUTOFMEMORY;
	}
	LOG_DEBUG(L"rateLimitedUIAEventHandler_create: done");
	return S_OK;
}

HRESULT rateLimitedUIAEventHandler_flush(RateLimitedEventHandler* pRateLimitedEventHandler) {
	LOG_DEBUG(L"rateLimitedUIAEventHandler_flush called");
	if(!pRateLimitedEventHandler) {
		LOG_DEBUG(L"rateLimitedUIAEventHandler_flush: invalid argument. Returning");
		return E_INVALIDARG;
	}
	pRateLimitedEventHandler->flush();
	return S_OK;
}

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
	EventRecord_t(const EventRecord_t& other) = delete;
	EventRecord_t& operator =(const EventRecord_t& other) = delete;
	CComPtr<IUIAutomationElement> element;
	bool isCoalesceable;
	std::vector<int> coalescingKey;
	unsigned int coalesceCount = 0;
	bool forceFlush;
	EventRecord_t(IUIAutomationElement* pElement, bool isCoAlesceable, bool forceFlush): element(pElement), isCoalesceable(isCoalesceable), forceFlush(forceFlush) {
		if(isCoalesceable) {
			coalescingKey = getRuntimeIDFromElement(element);
		}
	}
	bool canCoalesce(const EventRecord_t& other) CONST {
		return isCoalesceable && other.isCoalesceable && coalescingKey == other.coalescingKey;
	}
};

class FocusChangedEventRecord_t: public EventRecord_t {
	public:
	FocusChangedEventRecord_t(IUIAutomationElement* pElement): EventRecord_t(pElement, false, true) {
	}
};

class AutomationEventRecord_t: public EventRecord_t {
	public:
	EVENTID eventID;
	AutomationEventRecord_t(IUIAutomationElement* pElement, EVENTID eventID): EventRecord_t(pElement, true, false), eventID(eventID) {
		coalescingKey.push_back(eventID);
	}
};

class PropertyChangedEventRecord_t: public EventRecord_t {
	public:
	PROPERTYID propertyID;
	CComVariant value;
	PropertyChangedEventRecord_t(IUIAutomationElement* pElement, PROPERTYID propertyID, VARIANT value): EventRecord_t(pElement, true, false), propertyID(propertyID), value(value) { 
		coalescingKey.push_back(UIA_AutomationPropertyChangedEventId);
		coalescingKey.push_back(propertyID);
	}
};

using AnyEventRecord_t = std::variant<AutomationEventRecord_t, FocusChangedEventRecord_t, PropertyChangedEventRecord_t>;

class RateLimitedEventHandler : public IUIAutomationEventHandler, public IUIAutomationFocusChangedEventHandler, public IUIAutomationPropertyChangedEventHandler {
private:
	unsigned long m_refCount;
	CComQIPtr<IUIAutomationEventHandler> m_pExistingAutomationEventHandler;
	CComQIPtr<IUIAutomationFocusChangedEventHandler> m_pExistingFocusChangedEventHandler;
	CComQIPtr<IUIAutomationPropertyChangedEventHandler> m_pExistingPropertyChangedEventHandler;
	HWND m_messageWindow;
	UINT m_flushMessage;
	std::mutex mtx;
	std::list<AnyEventRecord_t> m_eventRecords;
	std::map<std::vector<int>, decltype(m_eventRecords)::iterator> m_eventRecordsByKey;

	template<typename EventRecordClass, typename... EventRecordArgTypes> HRESULT queueEvent(EventRecordArgTypes&&... args) {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent called");
		bool needsFlush = false;
		unsigned int flushTimeMS = 30;
		{ std::lock_guard lock(mtx);
			if(m_eventRecords.empty()) {
				LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: First event, needs callback.");
				needsFlush = true;
			}
			LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: Inserting new event");
			auto& recordVar = m_eventRecords.emplace_back(std::in_place_type_t<EventRecordClass>(),args...);
			auto recordVarIter = m_eventRecords.end();
			recordVarIter--;
			auto& record = std::get<EventRecordClass>(recordVar);
			if(record.forceFlush) {
				needsFlush = true;
				flushTimeMS = 0;
			}
			if(record.isCoalesceable) {
				LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: Is a coalesceable event");
				record.coalesceCount += 1;
				auto existingKeyIter = m_eventRecordsByKey.find(record.coalescingKey);
				if(existingKeyIter != m_eventRecordsByKey.end()) {
					LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: found existing event with same key"); 
					auto existingRecordVarIter = existingKeyIter->second;
					std::visit([&](EventRecord_t& existingRecord) {
						record.coalesceCount += existingRecord.coalesceCount;
					}, *existingRecordVarIter);
					LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: updating key");
					existingKeyIter->second = recordVarIter;
					LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: erasing old item"); 
					m_eventRecords.erase(existingRecordVarIter);
				} else {
					LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: Adding key");
					m_eventRecordsByKey.insert_or_assign(record.coalescingKey, recordVarIter);
				}
			}
		}
		if(needsFlush) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::queueEvent: posting flush message");
			PostMessage(m_messageWindow, m_flushMessage, reinterpret_cast<WPARAM>(this), flushTimeMS);
		}
		return S_OK;
	}

	HRESULT emitEvent(const AutomationEventRecord_t& record) const {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::emitAutomationEvent called");
		return m_pExistingAutomationEventHandler->HandleAutomationEvent(record.element, record.eventID);
	}

	HRESULT emitEvent(const FocusChangedEventRecord_t& record) const {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::emitFocusChangedEvent called");
		return m_pExistingFocusChangedEventHandler->HandleFocusChangedEvent(record.element);
	}

	HRESULT emitEvent(const PropertyChangedEventRecord_t& record) const {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::emitPropertyChangedEvent called");
		return m_pExistingPropertyChangedEventHandler->HandlePropertyChangedEvent(record.element, record.propertyID, record.value);
	}

	~RateLimitedEventHandler() {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::~RateLimitedUIAEventHandler called");
	}

public:

	RateLimitedEventHandler(IUnknown* pExistingHandler, HWND messageWindow, UINT flushMessage)
		: m_messageWindow(messageWindow), m_flushMessage(flushMessage), m_refCount(1), m_pExistingAutomationEventHandler(pExistingHandler), m_pExistingFocusChangedEventHandler(pExistingHandler), m_pExistingPropertyChangedEventHandler(pExistingHandler) {
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
		return queueEvent<AutomationEventRecord_t>(pElement, eventID);
	}

	// IUIAutomationFocusEventHandler method
	HRESULT STDMETHODCALLTYPE HandleFocusChangedEvent(IUIAutomationElement* pElement) {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::HandleFocusChangedEvent called");
		if(!m_pExistingFocusChangedEventHandler) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::HandleFocusChangedEvent: No existing focusChangeEventHandler, returning");
			return E_NOTIMPL;
		}
				return queueEvent<FocusChangedEventRecord_t>(pElement);
	}

	// IUIAutomationPropertyChangedEventHandler method
	HRESULT STDMETHODCALLTYPE HandlePropertyChangedEvent(
	IUIAutomationElement* pElement, PROPERTYID propertyID, VARIANT newValue) {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::HandlePropertyChangedEvent called");
		if(!m_pExistingPropertyChangedEventHandler) {
			LOG_DEBUG(L"RateLimitedUIAEventHandler::HandlePropertyChangedEvent: no existing handler. Returning");
			return E_NOTIMPL;
		}
		return queueEvent<PropertyChangedEventRecord_t>(pElement, propertyID, newValue);
	}

	void flush() {
		LOG_DEBUG(L"RateLimitedUIAEventHandler::flush called");
		decltype(m_eventRecords) eventRecordsCopy;
		decltype(m_eventRecordsByKey) eventRecordsByKeyCopy;
		{ std::lock_guard lock(mtx);
			eventRecordsCopy.swap(m_eventRecords);
			eventRecordsByKeyCopy.swap(m_eventRecordsByKey);
		}

		// Emit events
		LOG_DEBUG(L"RateLimitedUIAEventHandler::flush: Emitting events...");
		for(const auto& recordVar: eventRecordsCopy) {
			std::visit([this](const auto& record) {
				if(record.coalesceCount > 1) {
					// Beep(440 + (record.coalesceCount), 40);
				}
				this->emitEvent(record);
			}, recordVar);
		}
		LOG_DEBUG(L"RateLimitedUIAEventHandler::flush: done emitting events"); 
	}

};

HRESULT rateLimitedUIAEventHandler_create(IUnknown* pExistingHandler, HWND messageWindow, UINT flushMessage, RateLimitedEventHandler** ppRateLimitedEventHandler) {
	LOG_DEBUG(L"rateLimitedUIAEventHandler_create called");
	if (!pExistingHandler || !ppRateLimitedEventHandler) {
		LOG_DEBUG(L"rateLimitedUIAEventHandler_create: invalid arguments. Returning");
		return E_INVALIDARG;
	}

	// Create the RateLimitedEventHandler instance
	*ppRateLimitedEventHandler = new RateLimitedEventHandler(pExistingHandler, messageWindow, flushMessage);
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

# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2021-2024 NV Access Limited


from ctypes import byref, windll
from typing import (
	Iterable,
	Optional,
	Any,
	Generator,
	cast,
)
from comtypes import GUID
from comInterfaces import UIAutomationClient as UIA
import winVersion
from logHandler import log
from ._remoteOps import remoteAlgorithms
from ._remoteOps.remoteTypes import (
	RemoteExtensionTarget,
	RemoteInt,
	RemoteArray,
)
from ._remoteOps import operation
from ._remoteOps import remoteAPI
from ._remoteOps.lowLevel import (
	TextUnit,
	AttributeId,
	StyleId,
	PropertyId,
	AutomationIdentifierType
)


_isSupported: bool = False


def isSupported() -> bool:
	"""
	Returns whether UIA remote operations are supported on this version of Windows.
	"""
	return _isSupported


def initialize(doRemote: bool, UIAClient: UIA.IUIAutomation):
	"""
	Initializes UI Automation remote operations.
	The following parameters are deprecated and ignored:
	@param doRemote: true if code should be executed remotely, or false for locally.
	@param UIAClient: the current instance of the UI Automation client library running in NVDA.
	"""
	global _isSupported
	_isSupported = winVersion.getWinVer() >= winVersion.WIN11
	return True


def terminate():
	"""Terminates UIA remote operations support."""
	global _isSupported
	_isSupported = False


def msWord_getCustomAttributeValue(
	docElement: UIA.IUIAutomationElement,
	textRange: UIA.IUIAutomationTextRange,
	customAttribID: int,
) -> Optional[Any]:
	guid_msWord_extendedTextRangePattern = GUID("{93514122-FF04-4B2C-A4AD-4AB04587C129}")
	guid_msWord_getCustomAttributeValue = GUID("{081ACA91-32F2-46F0-9FB9-017038BC45F8}")
	op = operation.Operation()

	@op.buildFunction
	def code(ra: remoteAPI.RemoteAPI):
		remoteDocElement = ra.newElement(docElement)
		remoteTextRange = ra.newTextRange(textRange)
		remoteCustomAttribValue = ra.newVariant()
		extendedTextRangeIsSupported = remoteDocElement.isExtensionSupported(
			guid_msWord_extendedTextRangePattern,
		)
		with ra.ifBlock(extendedTextRangeIsSupported.inverse()):
			ra.logRuntimeMessage("docElement does not support extendedTextRangePattern")
			ra.Return(None)
		ra.logRuntimeMessage("docElement supports extendedTextRangePattern")
		remoteResult = ra.newVariant()
		ra.logRuntimeMessage("doing callExtension for extendedTextRangePattern")
		remoteDocElement.callExtension(
			guid_msWord_extendedTextRangePattern,
			remoteResult,
		)
		with ra.ifBlock(remoteResult.isNull()):
			ra.logRuntimeMessage("extendedTextRangePattern is null")
			ra.Return(None)
		ra.logRuntimeMessage("got extendedTextRangePattern")
		remoteExtendedTextRangePattern = remoteResult.asType(RemoteExtensionTarget)
		customAttributeValueIsSupported = remoteExtendedTextRangePattern.isExtensionSupported(
			guid_msWord_getCustomAttributeValue,
		)
		with ra.ifBlock(customAttributeValueIsSupported.inverse()):
			ra.logRuntimeMessage("extendedTextRangePattern does not support getCustomAttributeValue")
			ra.Return(None)
		ra.logRuntimeMessage("extendedTextRangePattern supports getCustomAttributeValue")
		ra.logRuntimeMessage("doing callExtension for getCustomAttributeValue")
		remoteExtendedTextRangePattern.callExtension(
			guid_msWord_getCustomAttributeValue,
			remoteTextRange,
			customAttribID,
			remoteCustomAttribValue,
		)
		ra.logRuntimeMessage("got customAttribValue of ", remoteCustomAttribValue)
		ra.Return(remoteCustomAttribValue)

	customAttribValue = op.execute()
	if customAttribValue is None:
		log.debugWarning("Custom attribute value not available")
		return None
	return customAttribValue


def collectAllHeadingsInTextRange(
	textRange: UIA.IUIAutomationTextRange,
) -> Generator[tuple[int, str, UIA.IUIAutomationElement], None, None]:
	op = operation.Operation()

	@op.buildIterableFunction
	def code(ra: remoteAPI.RemoteAPI):
		remoteTextRange = ra.newTextRange(textRange, static=True)
		with remoteAlgorithms.remote_forEachUnitInTextRange(
			ra,
			remoteTextRange,
			TextUnit.Paragraph,
		) as paragraphRange:
			val = paragraphRange.getAttributeValue(AttributeId.StyleId)
			with ra.ifBlock(val.isInt()):
				intVal = val.asType(RemoteInt)
				with ra.ifBlock((intVal >= StyleId.Heading1) & (intVal <= StyleId.Heading9)):
					level = (intVal - StyleId.Heading1) + 1
					label = paragraphRange.getText(-1)
					ra.Yield(level, label, paragraphRange)

	for level, label, paragraphRange in op.iterExecute(maxTries=20):
		yield level, label, paragraphRange


def findFirstHeadingInTextRange(
	textRange: UIA.IUIAutomationTextRange,
	wantedLevel: int | None = None,
	reverse: bool = False,
) -> tuple[int, str, UIA.IUIAutomationElement] | None:
	op = operation.Operation()

	@op.buildFunction
	def code(ra: remoteAPI.RemoteAPI):
		remoteTextRange = ra.newTextRange(textRange, static=True)
		remoteWantedLevel = ra.newInt(wantedLevel or 0)
		executionCount = ra.newInt(0, static=True)
		executionCount += 1
		ra.logRuntimeMessage("executionCount is ", executionCount)
		with ra.ifBlock(executionCount == 1):
			ra.logRuntimeMessage("Doing initial move")
			remoteTextRange.getLogicalAdapter(reverse).start.moveByUnit(TextUnit.Paragraph, 1)
		with remoteAlgorithms.remote_forEachUnitInTextRange(
			ra,
			remoteTextRange,
			TextUnit.Paragraph,
			reverse=reverse,
		) as paragraphRange:
			val = paragraphRange.getAttributeValue(AttributeId.StyleId)
			with ra.ifBlock(val.isInt()):
				intVal = val.asType(RemoteInt)
				with ra.ifBlock((intVal >= StyleId.Heading1) & (intVal <= StyleId.Heading9)):
					level = (intVal - StyleId.Heading1) + 1
					with ra.ifBlock((remoteWantedLevel == 0) | (level == remoteWantedLevel)):
						ra.logRuntimeMessage("found heading at level ", level)
						label = paragraphRange.getText(-1)
						ra.Return(level, label, paragraphRange)

	try:
		level, label, paragraphRange = op.execute(maxTries=20)
	except operation.NoReturnException:
		return None
	return (
		cast(int, level),
		cast(str, label),
		cast(UIA.IUIAutomationTextRange, paragraphRange),
	)

_propertyIdToGuidCache = {}
def propertyIdToGuid(propertyId: int) -> GUID:
	guid = _propertyIdToGuidCache.get(propertyId)
	if guid is not None:
		return guid
	guid = GUID()
	res = windll.UIAutomationCore.UiaLookupId(propertyId, byref(guid))
	if res != 0 or not guid:
		raise RuntimeError(f"Failed to convert propertyId {propertyId} to GUID")
	_propertyIdToGuidCache[propertyId] = guid
	return guid

def collectAllDataForTextRange(
	arg_textRange: UIA.IUIAutomationTextRange,
	arg_neededProperties: list[int],
	arg_neededCustomProperties: list[GUID],
	arg_neededAttributes: list[int],
	arg_rootElement: UIA.IUIAutomationElement,
	localMode: bool = False,
) -> Generator[tuple[str, list[Any], list[UIA.IUIAutomationElement]], None, None]:
	op = operation.Operation(enableCompiletimeLogging=False, localMode=localMode, enableRuntimeLogging=False)

	import time
	startTime = time.time()
	@op.buildIterableFunction
	def code(ra: remoteAPI.RemoteAPI):
		logicalFullTextRange = ra.newTextRange(arg_textRange).getLogicalAdapter()
		rootElement = ra.newElement(arg_rootElement)
		textPattern = rootElement.getTextPattern()
		rootElementId = rootElement.getPropertyValue(PropertyId.RuntimeId).stringify()
		cacheRequest = ra.newCacheRequest()
		for prop in arg_neededProperties:
			cacheRequest.addProperty(prop)
		for propGuid in arg_neededCustomProperties:
			cacheRequest.addCustomProperty(propGuid)
		elementMap = ra.newStringMap()
		ra.Yield(elementMap)
		numTotalElements = ra.newInt(0)
		numCachedElementsFetched = ra.newInt(0)
		walkingTextRange = ra.newTextRange(arg_textRange, static=True)
		with remoteAlgorithms.remote_forEachUnitInTextRange(
			ra,
			walkingTextRange,
			TextUnit.Format,
		) as formatRange:
			text = formatRange.getText(-1)
			logicalFormatRange = formatRange.getLogicalAdapter(reverse=False)
			logicalFormatRange.end = logicalFormatRange.start
			logicalFormatRange.textRange.expandToEnclosingUnit(TextUnit.Character)
			attributes = ra.newArray()
			for attributeId in arg_neededAttributes:
				val = formatRange.getAttributeValue(attributeId)
				if not localMode and attributeId == AttributeId.AnnotationTypes:
					newVals = ra.newArray()
					with ra.ifBlock(val.isArray()):
						with ra.forEachItemInArray(val.asType(RemoteArray)) as item:
							with ra.ifBlock(item.isInt()):
								guidVal = ra.lookupGuidFromAutomationIdentifier(item.asType(RemoteInt), AutomationIdentifierType.Annotation)
								newVals.append(guidVal)
							with ra.elseBlock():
								newVals.append(item)
					with ra.elseBlock():
						with ra.ifBlock(val.isInt()):
							guidVal = ra.lookupGuidFromAutomationIdentifier(val.asType(RemoteInt), AutomationIdentifierType.Annotation)
							newVals.append(guidVal)
						with ra.elseBlock():
							newVals.append(val)
					attributes.append(newVals)
				else:
					attributes.append(val)
			Element = formatRange.getEnclosingElement()
			elementId = Element.getPropertyValue(PropertyId.RuntimeId).stringify()
			deepestRemoteElementId = elementId.copy()
			with ra.whileBlock(lambda: Element.isNull().inverse()):
				numTotalElements += 1
				with ra.ifBlock(elementMap.hasKey(elementId)):
					numCachedElementsFetched += 1
					ra.breakLoop()
				Element.populateCache(cacheRequest)
				elementInfo = ra.newStringMap()
				elementMap[elementId] = elementInfo
				elementInfo['element'] = Element
				logicalChildRange = textPattern.rangeFromChild(Element).getLogicalAdapter()
				clippedStart = logicalChildRange.start < logicalFullTextRange.start
				elementInfo['clippedStart'] = clippedStart
				clippedEnd = logicalChildRange.end > logicalFullTextRange.end
				elementInfo['clippedEnd'] = clippedEnd
				with ra.ifBlock(elementId == rootElementId):
					ra.breakLoop()
				parentElement = Element.getParentElement()
				Element.set(parentElement)
				with ra.ifBlock(Element.isNull().inverse()):
					elementId.set(Element.getPropertyValue(PropertyId.RuntimeId).stringify())
					elementInfo['parentElementId'] = elementId
			ra.Yield(text, attributes, deepestRemoteElementId)
		ra.logRuntimeMessage("fetched ", numCachedElementsFetched, " elements out of ", numTotalElements)
	endTime = time.time()
	log.info(f"Building collectAllDataForTextRange took {endTime - startTime:.2f} seconds")

	import time
	startTime = time.time()
	results = list(op.iterExecute(maxTries=100))
	endTime = time.time()
	log.info(f"collectAllDataForTextRange took {endTime - startTime:.2f} seconds")
	elementMap = {}
	def resultsProcessor():
		for record in results:
			if isinstance(record, dict):
				elementMap.update(record)
			else:
				text, attribValues, deepestRemoteElementId = record
				ancestorIds = []
				elementId = deepestRemoteElementId
				while elementId:
					elementInfo = elementMap[elementId]
					ancestorIds.append(elementId)
					elementId = elementInfo.get("parentElementId")
				attribsMap = {arg_neededAttributes[i].value : attribValues[i] for i in range(len(arg_neededAttributes))}
				annotationTypesVal = attribsMap.get(AttributeId.AnnotationTypes.value)
				if isinstance(annotationTypesVal, GUID):
					ID = windll.UIAutomationCore.UiaLookupId(AutomationIdentifierType.Annotation, byref(annotationTypesVal))
					if ID != 0:
						annotationTypesVal = ID
					else:
						log.debugWarning(f"Failed to convert GUID {annotationTypesVal} to ID")
				elif isinstance(annotationTypesVal, Iterable):
					resolvedAnnotationTypes = []
					for item in annotationTypesVal:
						if isinstance(item, GUID):
							ID = windll.UIAutomationCore.UiaLookupId(AutomationIdentifierType.Annotation, byref(item))
							if ID != 0:
								resolvedAnnotationTypes.append(ID)
							else:
								log.debugWarning(f"Failed to convert GUID {item} to ID")
					annotationTypesVal = resolvedAnnotationTypes
				attribsMap[AttributeId.AnnotationTypes.value] = annotationTypesVal
				yield text, attribsMap, ancestorIds
	return elementMap, resultsProcessor()

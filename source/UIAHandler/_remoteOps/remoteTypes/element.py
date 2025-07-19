# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2023-2024 NV Access Limited


from __future__ import annotations
from typing import (
	cast,
)
from ctypes import (
	POINTER,
)
from comtypes import (
	GUID,
)
from UIAHandler import UIA
from .. import lowLevel
from .. import instructions
from ..remoteFuncWrapper import (
	remoteMethod,
)
from . import (
	RemoteExtensionTarget,
	RemoteIntEnum,
	RemoteBool,
	RemoteVariant,
	RemoteGuid,
)


class RemoteElement(RemoteExtensionTarget[POINTER(UIA.IUIAutomationElement)]):
	"""
	Represents a remote UI Automation element.
	Allows for navigation and property retrieval.
	"""

	_IsTypeInstruction = instructions.IsElement
	LocalType = POINTER(UIA.IUIAutomationElement)

	def _initOperand(self, initialValue: None = None, const: bool = False):
		if initialValue is not None:
			raise TypeError("Cannot initialize RemoteElement with an initial value.")
		return super()._initOperand()

	@property
	def localValue(self) -> UIA.IUIAutomationElement:
		value = super().localValue
		if value is None:
			return POINTER(UIA.IUIAutomationElement)()
		return cast(UIA.IUIAutomationElement, value.QueryInterface(UIA.IUIAutomationElement))

	@remoteMethod
	def getPropertyValue(
		self,
		propertyId: RemoteIntEnum[lowLevel.PropertyId] | lowLevel.PropertyId,
		ignoreDefault: RemoteBool | bool = False,
	) -> RemoteVariant:
		result = RemoteVariant(self.rob, self.rob.requestNewOperandId())
		self.rob.getDefaultInstructionList().addInstruction(
			instructions.ElementGetPropertyValue(
				result=result,
				target=self,
				propertyId=RemoteIntEnum.ensureRemote(self.rob, propertyId),
				ignoreDefault=RemoteBool.ensureRemote(self.rob, ignoreDefault),
			),
		)
		return result

	@remoteMethod
	def getCustomPropertyValue(
		self,
		propertyId: RemoteGuid | GUID,
		ignoreDefault: RemoteBool | bool = False,
	) -> RemoteVariant:
		result = RemoteVariant(self.rob, self.rob.requestNewOperandId())
		self.rob.getDefaultInstructionList().addInstruction(
			instructions.ElementGetPropertyValue(
				result=result,
				target=self,
				propertyId=RemoteGuid.ensureRemote(self.rob, propertyId).lookupId(lowLevel.AutomationIdentifierType.Property),
				ignoreDefault=RemoteBool.ensureRemote(self.rob, ignoreDefault),
			),
		)
		return result

	def _navigate(self, navigationDirection: lowLevel.NavigationDirection) -> RemoteElement:
		result = RemoteElement(self.rob, self.rob.requestNewOperandId())
		self.rob.getDefaultInstructionList().addInstruction(
			instructions.ElementNavigate(
				result=result,
				target=self,
				direction=RemoteIntEnum.ensureRemote(self.rob, navigationDirection),
			),
		)
		return result

	@remoteMethod
	def getParentElement(self) -> RemoteElement:
		return self._navigate(lowLevel.NavigationDirection.Parent)

	@remoteMethod
	def getFirstChildElement(self) -> RemoteElement:
		return self._navigate(lowLevel.NavigationDirection.FirstChild)

	@remoteMethod
	def getLastChildElement(self) -> RemoteElement:
		return self._navigate(lowLevel.NavigationDirection.LastChild)

	@remoteMethod
	def getNextSiblingElement(self) -> RemoteElement:
		return self._navigate(lowLevel.NavigationDirection.NextSibling)

	@remoteMethod
	def getPreviousSiblingElement(self) -> RemoteElement:
		return self._navigate(lowLevel.NavigationDirection.PreviousSibling)

	@remoteMethod
	def populateCache(self, request: RemoteCacheRequest) -> RemoteElement:
		self.rob.getDefaultInstructionList().addInstruction(
			instructions.ElementPopulateCache(
				target=self,
				request=request,
			),
		)
		return self

	@remoteMethod
	def getTextPattern(self) -> RemoteTextPattern:
		result = RemoteTextPattern(self.rob, self.rob.requestNewOperandId())
		self.rob.getDefaultInstructionList().addInstruction(
			instructions.ElementGetTextPattern(
				result=result,
				target=self,
			),
		)
		return result

from .textPattern import RemoteTextPattern

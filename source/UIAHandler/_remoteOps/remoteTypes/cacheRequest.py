# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2023-2024 NV Access Limited


from __future__ import annotations
from typing import (
	Iterable,
)
from comtypes import GUID
from comInterfaces import UIAutomationClient as UIA
from .. import lowLevel
from .. import instructions
from ..remoteFuncWrapper import (
	remoteMethod,
)
from . import (
	RemoteBaseObject,
	RemoteInt,
	RemoteGuid,
	RemoteIntEnum,
	RemoteVariant,
)


class RemoteCacheRequest(RemoteBaseObject[UIA.IUIAutomationCacheRequest]):
	"""
	Represents a cache request in the remote operation VM.
	"""

	def _generateInitInstructions(self) -> Iterable[instructions.InstructionBase]:
		yield instructions.NewCacheRequest(
			result=self,
		)

	@remoteMethod
	def isNull(self):
		variant = RemoteVariant(self.rob, self.operandId)
		return variant.isNull()

	@remoteMethod
	def addProperty(self, propertyId: RemoteIntEnum[lowLevel.PropertyId] | lowLevel.PropertyId):
		self.rob.getDefaultInstructionList().addInstruction(
			instructions.CacheRequestAddProperty(
				target=self,
				propertyId=RemoteIntEnum[lowLevel.PropertyId].ensureRemote(self.rob, propertyId)
			),
		)

	@remoteMethod
	def addCustomProperty(self, propertyId: RemoteGuid | GUID):
		self.rob.getDefaultInstructionList().addInstruction(
			instructions.CacheRequestAddProperty(
				target=self,
				propertyId=RemoteGuid.ensureRemote(self.rob, propertyId).lookupId(lowLevel.AutomationIdentifierType.Property)
			),
		)

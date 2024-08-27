# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2023-2024 NV Access Limited

"""
This module contains the instructions that operate on element cache requests.
Including the instructions that create and manipulate the cache request.
"""

from __future__ import annotations
from typing import cast
from dataclasses import dataclass
import UIAHandler
from comInterfaces import UIAutomationClient as UIA
from .. import lowLevel
from .. import builder
from ._base import _TypedInstruction


@dataclass
class NewCacheRequest(_TypedInstruction):
	opCode = lowLevel.InstructionType.NewCacheRequest
	result: builder.Operand

	def localExecute(self, registers: dict[lowLevel.OperandId, object]):
		if not UIAHandler.handler:
			raise RuntimeError("UIAHandler not initialized")
		client = cast(UIA.IUIAutomation, UIAHandler.handler.clientObject)
		registers[self.result.operandId] = client.CreateCacheRequest()


@dataclass
class IsCacheRequest(_TypedInstruction):
	opCode = lowLevel.InstructionType.IsCacheRequest
	result: builder.Operand
	target: builder.Operand

	def localExecute(self, registers: dict[lowLevel.OperandId, object]):
		registers[self.result.operandId] = isinstance(
			registers[self.target.operandId], UIA.IUIAutomationCacheRequest
		)


@dataclass
class CacheRequestAddProperty(_TypedInstruction):
	opCode = lowLevel.InstructionType.CacheRequestAddProperty
	target: builder.Operand
	propertyId: builder.Operand

	def localExecute(self, registers: dict[lowLevel.OperandId, object]):
		target = cast(UIA.IUIAutomationCacheRequest, registers[self.target.operandId])
		propertyId = cast(int, registers[self.propertyId.operandId])
		target.AddProperty(propertyId)

"""Microsoft Word 原生修订桥接。"""

from .bridge import WordBridge, WordBridgeError, RevisionInfo

__all__ = ["RevisionInfo", "WordBridge", "WordBridgeError"]

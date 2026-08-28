from unittest.mock import Mock

import pytest

from word_bridge.bridge import WordBridge


def test_revision_index_is_zero_based_for_com_collection():
    bridge = WordBridge(Mock(), Mock())
    bridge.document.Revisions.Count = 1
    revision = Mock()
    bridge.document.Revisions.Item.return_value = revision

    assert bridge._revision_at(0) is revision
    bridge.document.Revisions.Item.assert_called_once_with(1)


def test_revision_index_out_of_range():
    bridge = WordBridge(Mock(), Mock())
    bridge.document.Revisions.Count = 0

    with pytest.raises(IndexError):
        bridge._revision_at(0)


def test_replace_requires_non_empty_old_text():
    bridge = WordBridge(Mock(), Mock())

    with pytest.raises(ValueError):
        bridge.replace_text("", "new")

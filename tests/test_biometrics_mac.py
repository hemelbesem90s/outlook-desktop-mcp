"""
Tests for macOS Touch ID Biometric Security Gate in Outlook Desktop MCP.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_desktop_mcp.biometrics import (
    is_biometrics_disabled,
    verify_biometric,
    verify_biometric_gate,
)
from outlook_desktop_mcp.server_mac import (
    bridge,
    complete_task,
    create_calendar_event,
    create_event,
    create_meeting,
    delete_calendar_event,
    delete_email,
    delete_event,
    delete_message,
    delete_task,
    list_attachments,
    list_events,
    list_folders,
    list_tasks,
    send_email,
    update_event,
)


def test_is_biometrics_disabled_env():
    with patch.dict(os.environ, {"OUTLOOK_MCP_DISABLE_BIOMETRICS": "1"}):
        assert is_biometrics_disabled() is True
    with patch.dict(os.environ, {"OUTLOOK_MCP_DISABLE_BIOMETRICS": "true"}):
        assert is_biometrics_disabled() is True
    with patch.dict(os.environ, {"OUTLOOK_MCP_DISABLE_BIOMETRICS": "0"}):
        assert is_biometrics_disabled() is False
    with patch.dict(os.environ, {}, clear=True):
        assert is_biometrics_disabled() is False


def test_verify_biometric_bypass():
    with patch.dict(os.environ, {"OUTLOOK_MCP_DISABLE_BIOMETRICS": "1"}):
        assert verify_biometric("Test action") is True


def test_verify_biometric_swift_success():
    with patch.dict(os.environ, {"OUTLOOK_MCP_DISABLE_BIOMETRICS": "0"}):
        with patch("shutil.which", return_value="/usr/bin/swift"):
            mock_proc = MagicMock(returncode=0)
            with patch("subprocess.run", return_value=mock_proc) as mock_run:
                assert verify_biometric("Authorize send") is True
                assert mock_run.call_count == 1
                args = mock_run.call_args[0][0]
                assert args[0] == "/usr/bin/swift"
                assert args[3] == "Authorize send"


def test_verify_biometric_swift_denied_returns_false_without_dialog():
    with patch.dict(os.environ, {"OUTLOOK_MCP_DISABLE_BIOMETRICS": "0"}):
        with patch("shutil.which", return_value="/usr/bin/swift"):
            # exit code 1 means user cancelled or failed authentication
            mock_proc = MagicMock(returncode=1)
            with patch("subprocess.run", return_value=mock_proc) as mock_run:
                assert verify_biometric("Authorize send") is False
                # Must NOT call osascript fallback on user cancel
                assert mock_run.call_count == 1


def test_verify_biometric_swift_exit_2_falls_back_to_osascript():
    with patch.dict(os.environ, {"OUTLOOK_MCP_DISABLE_BIOMETRICS": "0"}):
        with patch("shutil.which", return_value="/usr/bin/swift"):
            # exit code 2 means policy cannot evaluate, trigger osascript fallback
            mock_swift = MagicMock(returncode=2)
            mock_dialog = MagicMock(returncode=0)
            with patch("subprocess.run", side_effect=[mock_swift, mock_dialog]) as mock_run:
                assert verify_biometric("Fallback prompt") is True
                assert mock_run.call_count == 2
                dialog_cmd = mock_run.call_args_list[1][0][0]
                assert dialog_cmd[0] == "osascript"
                assert "Fallback prompt" in dialog_cmd[2]


@pytest.mark.asyncio
async def test_mutating_tools_blocked_when_biometric_denied():
    with patch("outlook_desktop_mcp.server_mac.verify_biometric_gate", return_value=False):
        # 1. send_email
        res_email = await send_email(to="alice@example.com", subject="Test", body="Hello")
        assert "Biometric verification required" in res_email

        # 2. create_event
        res_evt = await create_event(
            subject="Test Event",
            start="2026-11-10T10:00:00",
            end="2026-11-10T10:30:00",
        )
        assert "Biometric verification required" in json.loads(res_evt)["error"]

        # 3. create_calendar_event alias
        res_cal_evt = await create_calendar_event(
            subject="Test Event",
            start="2026-11-10T10:00:00",
            end="2026-11-10T10:30:00",
        )
        assert "Biometric verification required" in json.loads(res_cal_evt)["error"]

        # 4. create_meeting
        res_meeting = await create_meeting(
            subject="Team Sync",
            start="2026-11-10T14:00:00",
            end="2026-11-10T15:00:00",
            required_attendees="team@example.com",
        )
        assert "Biometric verification required" in res_meeting

        # 5. update_event
        res_upd = await update_event(entry_id="42", location="Room A")
        assert "Biometric verification required" in json.loads(res_upd)["error"]

        # 6. delete_event
        res_del_evt = await delete_event(entry_id="42")
        assert "Biometric verification required" in res_del_evt

        # 7. delete_calendar_event alias
        res_del_cal_evt = await delete_calendar_event(entry_id="42")
        assert "Biometric verification required" in res_del_cal_evt

        # 8. delete_message
        res_del_msg = await delete_message(entry_id="42")
        assert "Biometric verification required" in res_del_msg

        # 9. delete_email alias
        res_del_eml = await delete_email(entry_id="42")
        assert "Biometric verification required" in res_del_eml

        # 10. delete_task
        res_del_task = await delete_task(entry_id="42")
        assert "Biometric verification required" in res_del_task


@pytest.mark.asyncio
async def test_update_event_attendee_confirmation_required():
    # Mock inspecting an event with 3 attendees
    inspect_meta = "42|||Team Sync|||organizer@example.com|||3|||Alice <a@ex.com>; Bob <b@ex.com>; Carol <c@ex.com>; |||false|||false"
    with patch.object(bridge, "run", return_value=inspect_meta):
        # 1. Without confirmation: should return confirmation_required and NOT call biometric gate
        gate_mock = MagicMock(return_value=True)
        with patch("outlook_desktop_mcp.server_mac.verify_biometric_gate", gate_mock):
            res = json.loads(await update_event(entry_id="42", location="Room B", confirm_send_updates=False))
            assert res["status"] == "confirmation_required"
            assert res["attendees_count"] == 3
            assert "Alice" in res["attendees"]
            assert gate_mock.call_count == 0

    # 2. With confirmation: proceeds to biometric gate and update script
    update_return = "42|||Team Sync|||Monday|||Tuesday|||Room B"
    with patch.object(bridge, "run", side_effect=[inspect_meta, update_return]):
        gate_mock = AsyncMock(return_value=True)
        with patch("outlook_desktop_mcp.server_mac.verify_biometric_gate", gate_mock):
            res = json.loads(await update_event(entry_id="42", location="Room B", confirm_send_updates=True))
            assert res["status"] == "updated"
            assert res["location"] == "Room B"
            assert res["attendees_notified"] is True
            assert gate_mock.call_count == 1
            prompt_used = gate_mock.call_args[0][0]
            assert "NOTIFYING 3 ATTENDEES" in prompt_used


@pytest.mark.asyncio
async def test_read_only_tools_do_not_invoke_biometric_gate():
    gate_mock = MagicMock(return_value=True)
    with patch("outlook_desktop_mcp.server_mac.verify_biometric_gate", gate_mock):
        with patch.object(
            bridge, "run", return_value="[]"
        ):
            await list_folders()
            await list_events()
            await list_tasks()
            # Gate should never have been called for read-only tools
            assert gate_mock.call_count == 0

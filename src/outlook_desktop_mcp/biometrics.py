"""
Biometric Security Gate for Outlook Desktop MCP.
================================================
Provides native macOS Touch ID verification for mutating operations
(sending emails, deleting messages, modifying/deleting calendar events)
via Apple's LocalAuthentication framework.
"""

from __future__ import annotations

import asyncio
import logging
import os
import shutil
import subprocess
import sys

logger = logging.getLogger("outlook_desktop_mcp.biometrics")

MACOS_SWIFT_TOUCHID = """
import Foundation
import LocalAuthentication

let context = LAContext()
var error: NSError?
let reason = CommandLine.arguments.count > 1 ? CommandLine.arguments[1] : "Outlook Desktop MCP Security Gate"

let policy: LAPolicy = context.canEvaluatePolicy(.deviceOwnerAuthenticationWithBiometrics, error: &error)
    ? .deviceOwnerAuthenticationWithBiometrics
    : .deviceOwnerAuthentication

if context.canEvaluatePolicy(policy, error: &error) {
    let semaphore = DispatchSemaphore(value: 0)
    var success = false
    context.evaluatePolicy(policy, localizedReason: reason) { (succ, _) in
        success = succ
        semaphore.signal()
    }
    _ = semaphore.wait(timeout: .distantFuture)
    exit(success ? 0 : 1)
} else {
    exit(2)
}
"""


def is_biometrics_disabled() -> bool:
    """Check if biometric security gate is bypassed via environment variable."""
    return os.environ.get("OUTLOOK_MCP_DISABLE_BIOMETRICS", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )


def verify_biometric(prompt_message: str, timeout_seconds: int = 60) -> bool:
    """Prompt for Touch ID on macOS before executing mutating operations.

    Args:
        prompt_message: Context message shown to the user on Touch ID prompt.
        timeout_seconds: Maximum time to wait for user verification (default 60s).

    Returns:
        True if verification succeeded or was bypassed in testing.
        False if user cancelled or verification failed.
    """
    if is_biometrics_disabled():
        logger.warning(
            "Biometric verification bypassed via OUTLOOK_MCP_DISABLE_BIOMETRICS"
        )
        return True

    if sys.platform != "darwin":
        # Non-macOS platforms do not use macOS Touch ID gate
        return True

    swift_bin = shutil.which("swift") or (
        "/usr/bin/swift" if os.path.exists("/usr/bin/swift") else None
    )
    if swift_bin:
        try:
            proc = subprocess.run(
                [swift_bin, "-e", MACOS_SWIFT_TOUCHID, prompt_message],
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
                check=False,
            )
            if proc.returncode == 0:
                logger.info("Biometric authorization approved: %s", prompt_message)
                return True
            if proc.returncode == 1:
                logger.warning(
                    "Biometric authorization denied or cancelled by user: %s",
                    prompt_message,
                )
                return False
            # Exit code 2 means policy cannot be evaluated (fallback needed)
            logger.info(
                "Swift Touch ID policy not evaluated (exit code 2), using osascript fallback"
            )
        except subprocess.TimeoutExpired:
            logger.error("Biometric verification timed out after %ds", timeout_seconds)
            return False
        except Exception as e:
            logger.warning(
                "Swift Touch ID execution error: %s, falling back to osascript dialog",
                e,
            )

    # osascript dialog fallback
    escaped = prompt_message.replace("\\", "\\\\").replace('"', '\\"')
    script = (
        f'display dialog "{escaped}" '
        'with title "Outlook Desktop MCP Security Gate" '
        'buttons {"Cancel", "Approve"} default button "Approve" '
        "with icon caution"
    )
    try:
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )
        if proc.returncode == 0:
            logger.info("Fallback dialog approved: %s", prompt_message)
            return True
        logger.warning("Fallback dialog cancelled by user: %s", prompt_message)
        return False
    except Exception as e:
        logger.error("Fallback dialog error: %s", e)
        return False


async def verify_biometric_gate(prompt_message: str, timeout_seconds: int = 60) -> bool:
    """Async wrapper for verify_biometric to avoid blocking the MCP event loop."""
    return await asyncio.to_thread(verify_biometric, prompt_message, timeout_seconds)

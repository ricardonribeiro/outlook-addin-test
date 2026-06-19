"""
Service Bus sender.

In Azure: the Function App's system-assigned identity has the
'Azure Service Bus Data Sender' role on the namespace (granted by Terraform).
DefaultAzureCredential resolves to Managed Identity — no connection string needed.

Locally: DefaultAzureCredential falls back to 'az login' credentials, which need
the 'Azure Service Bus Data Sender' role. If that role hasn't been granted, set
SERVICEBUS_CONNECTION_STRING in local.settings.json with the namespace's primary
connection string — that bypasses the RBAC requirement entirely.
"""

import hashlib
import json
import logging
import os

from azure.identity import DefaultAzureCredential
from azure.servicebus import ServiceBusClient, ServiceBusMessage

log = logging.getLogger(__name__)

# Module-level credential for reuse across invocations (token caching).
_credential = DefaultAzureCredential()


def send_to_queue(payload: dict, *, message_id: str) -> None:
    """
    Serialise payload to JSON and send it to the configured Service Bus queue.

    message_id is set as the Service Bus message ID so that the queue's
    duplicate detection can deduplicate re-sends of the same email.

    STUB: an app-level dedupe store (e.g. Redis / Cosmos) would go here for
    stronger guarantees across the duplicate detection window.
    """
    queue_name = os.environ["SERVICEBUS_QUEUE_NAME"]
    conn_str = os.getenv("SERVICEBUS_CONNECTION_STRING", "")

    if conn_str:
        client_ctx = ServiceBusClient.from_connection_string(conn_str)
    else:
        fqdn = os.environ["SERVICEBUS_FQDN"]
        client_ctx = ServiceBusClient(fqdn, credential=_credential)

    # Service Bus message IDs are capped at 128 chars; Outlook itemIds are 200+.
    # SHA-256 hex (64 chars) is collision-resistant and deterministic for the same email.
    sb_message_id = hashlib.sha256(message_id.encode()).hexdigest()

    with client_ctx as client, client.get_queue_sender(queue_name) as sender:
        msg = ServiceBusMessage(
            body=json.dumps(payload),
            message_id=sb_message_id,  # enables Service Bus duplicate detection
            content_type="application/json",
        )
        sender.send_messages(msg)
        log.info(
            "Message sent queue_name=%s message_id=%s", queue_name, message_id
        )

"""
JWT validation for tokens issued by Entra ID (Azure AD).

Tokens are acquired by the add-in via MSAL Nested App Authentication (NAA).
We validate signature, expiry, audience, and issuer. Both v1 and v2 issuers
are accepted; MSAL/NAA uses v2 (login.microsoftonline.com) in practice.
"""

import logging
import os
from typing import Any

import jwt
from jwt import PyJWKClient

log = logging.getLogger(__name__)

TENANT_ID = os.environ["TENANT_ID"]
API_AUDIENCE = os.environ["API_AUDIENCE"]

# Both v1 (sts.windows.net) and v2 (login.microsoftonline.com/v2.0) issuers are valid.
_VALID_ISSUERS = frozenset(
    [
        f"https://sts.windows.net/{TENANT_ID}/",
        f"https://login.microsoftonline.com/{TENANT_ID}/v2.0",
    ]
)

# PyJWKClient caches signing keys in memory; cache_keys=True is the default.
_jwks_client = PyJWKClient(
    f"https://login.microsoftonline.com/{TENANT_ID}/discovery/v2.0/keys",
    cache_keys=True,
)


class TokenValidationError(Exception):
    """Raised when the JWT cannot be validated; maps to HTTP 401."""


def validate_token(authorization_header: str) -> dict[str, Any]:
    """
    Validate the Bearer token from the Authorization header.

    Returns the decoded JWT claims dict on success.
    Raises TokenValidationError on any validation failure (caller returns 401).
    """
    if not authorization_header or not authorization_header.startswith("Bearer "):
        raise TokenValidationError(
            "Missing or malformed Authorization header — expected 'Bearer <token>'"
        )

    token = authorization_header[7:]  # strip "Bearer "

    try:
        signing_key = _jwks_client.get_signing_key_from_jwt(token)
    except Exception as exc:
        raise TokenValidationError(f"Failed to resolve signing key: {exc}") from exc

    try:
        # Decode and verify signature + expiry + audience.
        # Issuer validation is done manually below to support both v1 and v2.
        claims: dict[str, Any] = jwt.decode(
            token,
            signing_key.key,
            algorithms=["RS256"],
            audience=API_AUDIENCE,
            options={
                "require": ["exp", "iss", "aud"],
                "verify_iss": False,  # we validate issuer manually (v1 + v2 support)
            },
        )
    except jwt.ExpiredSignatureError as exc:
        raise TokenValidationError("Token has expired") from exc
    except jwt.InvalidAudienceError as exc:
        raise TokenValidationError(f"Invalid token audience — expected {API_AUDIENCE!r}") from exc
    except jwt.PyJWTError as exc:
        raise TokenValidationError(f"Token validation failed: {exc}") from exc

    issuer = claims.get("iss", "")
    if issuer not in _VALID_ISSUERS:
        raise TokenValidationError(
            f"Unexpected token issuer: {issuer!r}. "
            f"Expected one of the v1/v2 issuers for tenant {TENANT_ID}"
        )

    log.debug("Token validated for upn=%s", claims.get("preferred_username") or claims.get("upn"))
    return claims

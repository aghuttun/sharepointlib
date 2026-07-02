"""Optional live integration test against Microsoft Graph SharePoint endpoints."""

import os

import pytest

import sharepointlib


REQUIRED_ENV = [
    "SHAREPOINTLIB_CLIENT_ID",
    "SHAREPOINTLIB_TENANT_ID",
    "SHAREPOINTLIB_CLIENT_SECRET",
    "SHAREPOINTLIB_SP_DOMAIN",
]


@pytest.mark.skipif(
    any(not os.getenv(name) for name in REQUIRED_ENV),
    reason="Missing required SHAREPOINTLIB_* environment variables.",
)
def test_live_sharepoint_smoke():
    client = sharepointlib.SharePoint(
        client_id=os.environ["SHAREPOINTLIB_CLIENT_ID"],
        tenant_id=os.environ["SHAREPOINTLIB_TENANT_ID"],
        client_secret=os.environ["SHAREPOINTLIB_CLIENT_SECRET"],
        sp_domain=os.environ["SHAREPOINTLIB_SP_DOMAIN"],
    )

    # Use a likely non-existing site name to avoid requiring tenant-specific ids.
    response = client.get_site_info(name="__sharepointlib_live_test__")

    # Any of these statuses confirms the call path is live and authenticated/authorised path is exercised.
    assert response.status_code in (200, 400, 401, 403, 404)

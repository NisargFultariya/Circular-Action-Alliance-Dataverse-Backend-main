"""
Dataverse Web API helpers — table/column creation.

Direct Python port of the Node.js POC's server.js logic (getToken, dvFetch,
buildFieldAttribute, handleCreate). Same request shapes, same behavior —
only the language changed.
"""

import json
import os
import re
import time
import httpx

# Metadata operations (create table/column, publish) can genuinely take well
# over 30s on Dataverse, especially right after a security-role change or on
# a busy environment — don't time these out too aggressively.
DEFAULT_TIMEOUT = 120.0
PUBLISH_TIMEOUT = 180.0
TOKEN_TIMEOUT = 20.0


class DataverseError(Exception):
    """Raised for any Dataverse/token call failure. `steps` carries the
    progress log accumulated before the failure, for the caller to report."""

    def __init__(self, message: str, steps: list | None = None):
        super().__init__(message)
        self.steps = steps or []


def load_config() -> dict | None:
    """Config comes ONLY from environment variables — never a committed file.
    Locally, use a `.env` file (see `.env.example`); on Render, set these as
    real Environment Variables in the service's dashboard.

    The first 4 are required (no sensible default exists for credentials).
    Everything below that was previously hardcoded in this file — each now
    has a fallback equal to what used to be baked in, so nothing breaks if
    you don't touch .env, but all of it can now be changed without editing
    code:
      DEFAULT_TABLE_PREFIX     — schema-name prefix for a table/field when
                                  the caller doesn't specify one.
      SOLUTION_UNIQUE_NAME     — unmanaged solution new tables/columns are
                                  added to. Empty (default) = the
                                  environment's Default solution.
      POWER_PAGES_WEBSITE_ID   — the mspp_website record Web API site
                                  settings/table permissions attach to.
      ADMINISTRATORS_WEBROLE_ID — the mspp_webrole bound to every table
                                  permission ensure_webapi_access creates.
      DEFAULT_OWNERSHIP_TYPE   — 'UserOwned' or 'OrganizationOwned' for a
                                  newly created table.
    """
    cfg = {
        "envUrl": os.environ.get("DATAVERSE_ENV_URL", "").strip().rstrip("/"),
        "tenantId": os.environ.get("DATAVERSE_TENANT_ID", "").strip(),
        "clientId": os.environ.get("DATAVERSE_CLIENT_ID", "").strip(),
        "clientSecret": os.environ.get("DATAVERSE_CLIENT_SECRET", "").strip(),
    }
    if not all(cfg.values()):
        return None

    cfg["defaultPrefix"] = os.environ.get("DEFAULT_TABLE_PREFIX", "new").strip() or "new"
    cfg["solutionUniqueName"] = os.environ.get("SOLUTION_UNIQUE_NAME", "").strip()
    cfg["websiteId"] = (
        os.environ.get("POWER_PAGES_WEBSITE_ID", "").strip()
        or "0acaa074-b128-4c0a-b47b-727352d5f03f"
    )
    cfg["adminWebRoleId"] = (
        os.environ.get("ADMINISTRATORS_WEBROLE_ID", "").strip()
        or "e19d8193-9d33-4ef7-b374-17493787636b"
    )
    cfg["defaultOwnershipType"] = (
        os.environ.get("DEFAULT_OWNERSHIP_TYPE", "").strip() or "UserOwned"
    )
    return cfg


def pascal_case(text: str) -> str:
    words = [w for w in re.split(r"[^a-zA-Z0-9]+", text or "") if w]
    return "".join(w[:1].upper() + w[1:] for w in words)


def label(text: str) -> dict:
    return {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        "LocalizedLabels": [
            {
                "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                "Label": text,
                "LanguageCode": 1033,
            }
        ],
    }


def required_level(value: str = "None") -> dict:
    return {
        "Value": value,
        "CanBeChanged": True,
        "ManagedPropertyLogicalName": "canmodifyrequirementlevelsettings",
    }


def _picklist_option_value(raw, index: int) -> int:
    """The SPA's OptionsEditor always sends a text `value` equal to the
    option's label (e.g. "Red", "option-1") — never a real numeric picklist
    value. Only trust `raw` as the Value when it actually parses as an int;
    otherwise fall back to the same auto-numbered scheme used when no value
    is sent at all, instead of letting int() raise and fail the whole
    table-creation call."""
    if raw not in (None, ""):
        try:
            return int(raw)
        except (TypeError, ValueError):
            pass
    return 100000000 + index


def _field_opts(f: dict) -> dict:
    """Builds the `opts` dict `build_field_attribute` expects from one field
    payload dict (the shape the SPA sends per-field, whether creating a new
    table or adding to an existing one) — every optional, type-specific
    property, only set when the SPA actually provided it. Shared by
    create_table_with_fields and add_fields_to_table so the two never drift."""
    return {
        "options": f.get("options"),
        "dateFormat": f.get("dateFormat"),
        # Only set when provided — build_field_attribute's own
        # opts.get(..., default) fallback needs the key absent (not
        # present-but-None) to apply its default range.
        **({"minValue": f["minValue"]} if f.get("minValue") is not None else {}),
        **({"maxValue": f["maxValue"]} if f.get("maxValue") is not None else {}),
        **({"precision": f["precision"]} if f.get("precision") is not None else {}),
        "allowMultiple": bool(f.get("allowMultiple")),
        "defaultValue": bool(f.get("defaultValue")),
        **(
            {"defaultOptionValue": f["defaultOptionValue"]}
            if f.get("defaultOptionValue")
            else {}
        ),
        **({"trueLabel": f["trueLabel"]} if f.get("trueLabel") else {}),
        **({"falseLabel": f["falseLabel"]} if f.get("falseLabel") else {}),
        **({"maxSizeKb": f["maxSizeKb"]} if f.get("maxSizeKb") is not None else {}),
        # Wires the "Required" checkbox into the real column for the first
        # time — build_field_attribute already reads this opt, it just never
        # received anything but the "None" default before now.
        **({"requiredLevel": "ApplicationRequired"} if f.get("required") else {}),
    }


def build_field_attribute(
    schema_name: str,
    display_name: str,
    description: str | None,
    field_type: str,
    opts: dict | None = None,
) -> dict:
    opts = opts or {}
    base = {
        "DisplayName": label(display_name),
        "Description": label(description or display_name),
        "RequiredLevel": required_level(opts.get("requiredLevel", "None")),
        "SchemaName": schema_name,
    }

    if field_type == "Text":  # single line of text
        return {
            **base,
            "AttributeType": "String",
            "AttributeTypeName": {"Value": "StringType"},
            "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
            "FormatName": {"Value": "Text"},
            "MaxLength": int(opts.get("maxLength") or 100),
        }

    if field_type == "Memo":  # multiline text
        return {
            **base,
            "AttributeType": "Memo",
            "AttributeTypeName": {"Value": "MemoType"},
            "@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
            "Format": "TextArea",
            "MaxLength": int(opts.get("maxLength") or 2000),
        }

    if field_type == "Integer":  # whole number
        return {
            **base,
            "AttributeType": "Integer",
            "AttributeTypeName": {"Value": "IntegerType"},
            "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
            "Format": "None",
            "MinValue": int(opts.get("minValue", -2147483648)),
            "MaxValue": int(opts.get("maxValue", 2147483647)),
        }

    if field_type == "Decimal":
        return {
            **base,
            "AttributeType": "Decimal",
            "AttributeTypeName": {"Value": "DecimalType"},
            "@odata.type": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
            "Precision": int(opts.get("precision", 2)),
            "MinValue": float(opts.get("minValue", -100000000000)),
            "MaxValue": float(opts.get("maxValue", 100000000000)),
        }

    if field_type == "Currency":
        return {
            **base,
            "AttributeType": "Money",
            "AttributeTypeName": {"Value": "MoneyType"},
            "@odata.type": "Microsoft.Dynamics.CRM.MoneyAttributeMetadata",
            "Precision": int(opts.get("precision", 2)),
            "MinValue": float(opts.get("minValue", 0)),
            "MaxValue": float(opts.get("maxValue", 100000000000)),
        }

    if field_type == "Boolean":  # yes/no
        return {
            **base,
            "AttributeType": "Boolean",
            "AttributeTypeName": {"Value": "BooleanType"},
            "@odata.type": "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
            "DefaultValue": bool(opts.get("defaultValue", False)),
            "OptionSet": {
                "OptionSetType": "Boolean",
                "TrueOption": {"Value": 1, "Label": label(opts.get("trueLabel", "Yes"))},
                "FalseOption": {"Value": 0, "Label": label(opts.get("falseLabel", "No"))},
            },
        }

    if field_type == "DateTime":  # date-only OR date-and-time, via Format
        return {
            **base,
            "AttributeType": "DateTime",
            "AttributeTypeName": {"Value": "DateTimeType"},
            "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
            # 'DateOnly' or 'DateAndTime' — the SPA's Date vs DateTime field
            # types both map here, distinguished only by this Format.
            "Format": opts.get("dateFormat", "DateOnly"),
        }

    if field_type == "Choice":  # picklist (or multi-select)
        options = [o for o in (opts.get("options") or []) if o and o.get("label")]
        if not options:
            raise ValueError("Choice column needs at least one option (label + value).")
        option_set = {
            "@odata.type": "Microsoft.Dynamics.CRM.OptionSetMetadata",
            "IsGlobal": False,
            "OptionSetType": "Picklist",
            "Options": [
                {
                    "Value": _picklist_option_value(o.get("value"), i),
                    "Label": label(o["label"]),
                }
                for i, o in enumerate(options)
            ],
        }
        if opts.get("allowMultiple"):
            return {
                **base,
                "AttributeType": "Virtual",
                "AttributeTypeName": {"Value": "MultiSelectPicklistType"},
                "@odata.type": "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata",
                "OptionSet": option_set,
            }
        picklist_attr = {
            **base,
            "AttributeType": "Picklist",
            "AttributeTypeName": {"Value": "PicklistType"},
            "@odata.type": "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
            "OptionSet": option_set,
        }
        # The option the admin picked as "Default value" — resolved through
        # the same numbering _picklist_option_value used to build the option
        # set above, so it always matches one of the real Values. Single-
        # select only; a multi-select picklist has no single DefaultFormValue.
        default_raw = opts.get("defaultOptionValue")
        if default_raw not in (None, ""):
            picklist_attr["DefaultFormValue"] = _picklist_option_value(default_raw, 0)
        return picklist_attr

    if field_type in ("File", "Image"):
        # File/Image columns attach to a row (Dataverse "File"/"Image" column
        # types) — same metadata shape, Image additionally supports
        # thumbnails. Confirmed live: "File" is NOT a valid AttributeType
        # enum member ("Requested value 'File' was not found.") — both use
        # "Virtual" here, same as Image always did; only AttributeTypeName
        # and @odata.type actually distinguish the two.
        return {
            **base,
            "AttributeType": "Virtual",
            "AttributeTypeName": {
                "Value": "FileType" if field_type == "File" else "ImageType"
            },
            "@odata.type": (
                "Microsoft.Dynamics.CRM.FileAttributeMetadata"
                if field_type == "File"
                else "Microsoft.Dynamics.CRM.ImageAttributeMetadata"
            ),
            "MaxSizeInKB": int(opts.get("maxSizeKb", 10240)),
        }

    if field_type == "Lookup":
        # A real Dataverse Lookup column needs a *relationship*, created via a
        # separate API (OneToManyRelationships) after both tables exist — it
        # can't be created as a plain attribute like the types above. Wire
        # this up once the "which table does it point to" UI exists.
        raise ValueError(
            "Lookup columns require a relationship, not a plain attribute — "
            "not implemented in this POC yet."
        )

    raise ValueError(f"Unsupported field type: {field_type}")


def get_token(cfg: dict) -> str:
    token_url = f"https://login.microsoftonline.com/{cfg['tenantId']}/oauth2/v2.0/token"
    body = {
        "client_id": cfg["clientId"],
        "client_secret": cfg["clientSecret"],
        "grant_type": "client_credentials",
        "scope": f"{cfg['envUrl']}/.default",
    }
    try:
        res = httpx.post(token_url, data=body, timeout=TOKEN_TIMEOUT)
    except httpx.HTTPError as err:
        raise DataverseError(f"Could not reach Microsoft's token endpoint: {err}") from err

    data = res.json() if res.content else {}
    if res.status_code >= 400:
        raise DataverseError(
            data.get("error_description") or data.get("error") or f"Token request failed ({res.status_code})"
        )
    token = data.get("access_token")
    if not token:
        raise DataverseError(
            "Token endpoint returned no access_token — check tenant ID / client ID / client secret."
        )
    return token


def dv_request(method: str, url: str, token: str, timeout: float = DEFAULT_TIMEOUT, **kwargs) -> httpx.Response:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        **({"Content-Type": "application/json; charset=utf-8"} if kwargs.get("json") is not None else {}),
        **kwargs.pop("headers", {}),
    }
    try:
        return httpx.request(method, url, headers=headers, timeout=timeout, **kwargs)
    except httpx.TimeoutException as err:
        raise DataverseError(
            f"{url} did not respond within {int(timeout)}s. Dataverse metadata operations can be "
            f"slow right after a security-role change — wait a minute and try again."
        ) from err
    except httpx.HTTPError as err:
        raise DataverseError(f"Could not reach {url}: {err}") from err


def create_table_and_field(payload: dict) -> dict:
    """Same flow as the Node POC's handleCreate(): token -> create table
    (with a primary column, and optionally one more field) -> publish."""
    steps: list[dict] = []
    cfg = load_config()
    if not cfg:
        raise DataverseError(
            "Backend is missing DATAVERSE_ENV_URL / DATAVERSE_TENANT_ID / "
            "DATAVERSE_CLIENT_ID / DATAVERSE_CLIENT_SECRET environment variables."
        )

    prefix = (
        re.sub(r"[^a-zA-Z0-9]", "", str(payload.get("prefixRaw") or cfg["defaultPrefix"])).lower()
        or cfg["defaultPrefix"]
    )

    steps.append({"label": "Requesting access token from Microsoft Entra ID", "status": "running"})
    token = get_token(cfg)
    steps[-1]["status"] = "done"

    table_display_name = payload["tableDisplayName"]
    primary_display_name = payload.get("primaryDisplayName") or "Name"
    table_schema_name = f"{prefix}_{pascal_case(table_display_name)}"
    primary_schema_name = f"{prefix}_{pascal_case(primary_display_name)}"

    attributes = [
        {
            "AttributeType": "String",
            "AttributeTypeName": {"Value": "StringType"},
            "Description": label(f"Primary name column for {table_display_name}"),
            "DisplayName": label(primary_display_name),
            "IsPrimaryName": True,
            "RequiredLevel": required_level("None"),
            "SchemaName": primary_schema_name,
            "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
            "FormatName": {"Value": "Text"},
            "MaxLength": int(payload.get("primaryMaxLength") or 100),
        }
    ]

    field_schema_name = None
    if payload.get("addField") and payload.get("fieldDisplayName") and payload.get("fieldType"):
        field_schema_name = f"{prefix}_{pascal_case(payload['fieldDisplayName'])}"
        attributes.append(
            build_field_attribute(
                field_schema_name,
                payload["fieldDisplayName"],
                payload.get("fieldDescription"),
                payload["fieldType"],
                payload.get("fieldOptions") or {},
            )
        )

    entity_body = {
        "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata",
        "SchemaName": table_schema_name,
        "Description": label(payload.get("tableDescription") or f"Dummy POC table: {table_display_name}"),
        "DisplayName": label(table_display_name),
        "DisplayCollectionName": label(payload.get("tableCollectionName") or f"{table_display_name}s"),
        "HasActivities": False,
        "HasNotes": False,
        "IsActivity": False,
        "OwnershipType": payload.get("ownershipType") or cfg["defaultOwnershipType"],
        "Attributes": attributes,
    }

    steps.append({"label": f'Creating table "{table_display_name}" ({table_schema_name})', "status": "running"})
    create_headers = {}
    solution_name = payload.get("solutionUniqueName") or cfg["solutionUniqueName"]
    if solution_name:
        create_headers["MSCRM.SolutionUniqueName"] = solution_name

    create_res = dv_request(
        "POST",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions",
        token,
        headers=create_headers,
        json=entity_body,
    )
    if create_res.status_code >= 400:
        steps[-1]["status"] = "failed"
        err_json = create_res.json() if create_res.content else {}
        message = (err_json.get("error") or {}).get("message") or f"Table creation failed ({create_res.status_code})"
        raise DataverseError(message, steps)
    steps[-1]["status"] = "done"

    entity_id_header = create_res.headers.get("odata-entityid", "")
    metadata_id_match = re.search(r"\(([0-9a-fA-F-]{36})\)", entity_id_header)
    metadata_id = metadata_id_match.group(1) if metadata_id_match else None
    logical_name = table_schema_name.lower()

    if payload.get("publish"):
        steps.append({"label": "Publishing customizations", "status": "running"})
        pub_res = dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/PublishAllXml",
            token,
            json={},
            timeout=PUBLISH_TIMEOUT,
        )
        steps[-1]["status"] = "done" if pub_res.status_code < 400 else "failed"
        if pub_res.status_code >= 400:
            err_json = pub_res.json() if pub_res.content else {}
            steps[-1]["note"] = (err_json.get("error") or {}).get("message") or f"Publish failed ({pub_res.status_code})"

    return {
        "success": True,
        "steps": steps,
        "table": {
            "displayName": table_display_name,
            "schemaName": table_schema_name,
            "logicalName": logical_name,
            "metadataId": metadata_id,
            "primaryColumn": {
                "displayName": primary_display_name,
                "schemaName": primary_schema_name,
                "logicalName": primary_schema_name.lower(),
            },
            "field": (
                {
                    "displayName": payload["fieldDisplayName"],
                    "schemaName": field_schema_name,
                    "logicalName": field_schema_name.lower(),
                    "type": payload["fieldType"],
                }
                if field_schema_name
                else None
            ),
            "envUrl": cfg["envUrl"],
        },
    }


# Entity sets for the two admin-facing catalog tables this backend writes
# results back into once a real table/column is created — see
# create_table_with_fields().
CONFIG_ENTITY_SET = "simform_reportconfigurations"
FIELDS_ENTITY_SET = "simform_reportfieldses"

# The Service Provider "Activity Reports" summary/header table (see
# sp-activity-reports.ts) and the hidden Lookup every template's own main
# table gets, pointing back at it — see _ensure_activity_report_lookup.
ACTIVITY_REPORTS_TABLE = "simform_serviceprovideractivityreports"
ACTIVITY_REPORT_LOOKUP_LOGICAL_NAME = "simform_activityreport"

# simform_displaystatus picklist, confirmed live: 1=Active, 2=Inactive.
DISPLAY_STATUS_ACTIVE = 1

# mspp_scope picklist value for "Global" — a fixed Power Pages platform
# constant (not per-environment config), unlike the ids below which now
# live in load_config()/`.env` (POWER_PAGES_WEBSITE_ID, ADMINISTRATORS_WEBROLE_ID).
GLOBAL_SCOPE = 756150000


def _upsert_site_setting(cfg: dict, token: str, name: str, value: str) -> None:
    """Create-or-update one Web API site setting by name (idempotent — safe
    to call every time a template is built, not just the first).

    Deletes EVERY existing record for this name, not just the first one the
    query happens to return — confirmed live this actually matters: an
    earlier version here only ever looked at `value[0]`, and a duplicate
    left behind by some other call path (e.g. ensure_webapi_access's own
    read-merge-write racing a concurrent one for a table several templates'
    Grid Table Lookups share, like the F&O released-products table) could
    sit there indefinitely — and since `$filter` with no `$orderby` returns
    rows in an unspecified order, a LATER read (this function's own, or
    `_site_setting_value`'s) could just as easily pick up the stale
    duplicate instead of the correct record, silently narrowing an
    already-fixed field allow-list back down. Deleting the whole set before
    recreating means there is only ever one record for this name once this
    function returns, regardless of how many accumulated before it ran."""
    url = (
        f"{cfg['envUrl']}/api/data/v9.2/mspp_sitesettings"
        f"?$filter=mspp_name eq '{name}' and _mspp_websiteid_value eq {cfg['websiteId']}"
        f"&$select=mspp_sitesettingid"
    )
    res = dv_request("GET", url, token, timeout=30.0)
    if res.status_code >= 400:
        raise DataverseError(f"Could not look up site setting '{name}' ({res.status_code})")
    existing_rows = res.json().get("value") or []

    # Confirmed live: PATCHing an existing mspp_sitesettings record fails
    # with "The given key was not present in the dictionary" (0x80040224) —
    # reproduced on multiple unrelated records, so it's an environment/
    # platform quirk, not something specific to one row. Delete + recreate
    # routes around it entirely; POST-ing a brand-new row always works.
    for row in existing_rows:
        delete_res = dv_request(
            "DELETE",
            f"{cfg['envUrl']}/api/data/v9.2/mspp_sitesettings({row['mspp_sitesettingid']})",
            token,
            timeout=30.0,
        )
        if delete_res.status_code >= 400:
            raise DataverseError(
                f"Could not replace site setting '{name}' — delete failed "
                f"({delete_res.status_code})"
            )

    body = {
        "mspp_name": name,
        "mspp_value": value,
        "mspp_websiteid@odata.bind": f"/mspp_websites({cfg['websiteId']})",
    }
    create_res = dv_request(
        "POST",
        f"{cfg['envUrl']}/api/data/v9.2/mspp_sitesettings",
        token,
        json=body,
        timeout=30.0,
    )
    if create_res.status_code >= 400:
        raise DataverseError(
            f"Could not create site setting '{name}' ({create_res.status_code}): "
            f"{create_res.text[:300]}"
        )


def _site_setting_value(cfg: dict, token: str, name: str) -> str | None:
    """The current value of one Web API site setting, or None if it doesn't
    exist yet.

    If more than one record for this name somehow still exists (a
    duplicate from before _upsert_site_setting cleaned up every match, or
    one that appeared between two calls), prefers the LONGEST comma-list
    value rather than an arbitrary/first one — for the `.../fields`
    settings this function is actually used for, the narrower duplicate is
    the stale one nearly every time (a caller that only ever asks to ADD
    fields, never remove them — see ensure_webapi_access's own merge —
    means the correct current state is always the union, i.e. the longest
    one seen)."""
    url = (
        f"{cfg['envUrl']}/api/data/v9.2/mspp_sitesettings"
        f"?$filter=mspp_name eq '{name}' and _mspp_websiteid_value eq {cfg['websiteId']}"
        f"&$select=mspp_value"
    )
    res = dv_request("GET", url, token, timeout=30.0)
    if res.status_code >= 400:
        raise DataverseError(f"Could not look up site setting '{name}' ({res.status_code})")
    rows = res.json().get("value") or []
    if not rows:
        return None
    return max((r.get("mspp_value") or "" for r in rows), key=len)


def _lookup_webrole_id(cfg: dict, token: str, name: str) -> str | None:
    """Finds a web role by name on this site — e.g. "Authenticated Users",
    the closest thing this site has to a generic "any signed-in contact"
    role (it has no dedicated Service Provider/Producer web roles at all).
    Looked up by name at call time rather than hardcoded, since a role's
    GUID is org-specific. Best-effort: None if not found, so a caller can
    skip granting it rather than fail the whole flow."""
    res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/mspp_webroles"
        f"?$select=mspp_webroleid&$filter=mspp_name eq '{name}' "
        f"and _mspp_websiteid_value eq {cfg['websiteId']}",
        token,
        timeout=30.0,
    )
    if res.status_code >= 400:
        return None
    match = (res.json().get("value") or [None])[0]
    return match.get("mspp_webroleid") if match else None


def _current_webrole_ids_for_permission(
    cfg: dict, token: str, permission_id: str
) -> list[str]:
    """The web roles actually bound to one entity permission — read from the
    classic N:N intersect table directly (`mspp_entitypermission_webroleset`),
    never via `$expand` on `mspp_entitypermission_webrole`: confirmed live,
    that $expand (and the Web Roles subgrid on the record's own form, which
    uses the same path) reads back empty even when the association genuinely
    exists. Querying the intersect table's own entity set is the only read
    path confirmed reliable."""
    res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/mspp_entitypermission_webroleset",
        token,
        params={"$filter": f"mspp_entitypermissionid eq {permission_id}"},
        timeout=30.0,
    )
    if res.status_code >= 400:
        return []
    return [row["mspp_webroleid"] for row in res.json().get("value", [])]


def _sync_powerpagecomponent_webroles(
    cfg: dict,
    token: str,
    permission_id: str,
    webrole_ids: list[str],
    entity_logical_name: str,
    entity_display_name: str,
    crud: dict,
) -> None:
    """Mirrors a table permission's web roles into its `powerpagecomponents`
    twin — REQUIRED on a site using Power Pages' Enhanced Data Model
    (confirmed live: this org has it enabled). Under the Enhanced Data
    Model, both the design Studio AND the live site's own runtime
    permission enforcement read a table permission from the JSON `content`
    column on `powerpagecomponents` (component type 18), specifically its
    `adx_entitypermission_webrole` array — NOT from the classic
    `mspp_entitypermission_webrole` N:N relationship this module otherwise
    writes through. Dataverse mirrors an entity permission's plain CRUD
    flags from the classic table into this JSON quickly, but does NOT
    reliably mirror the web role array — confirmed live, a freshly-bound
    classic association can sit for a very long time (if ever) without
    appearing here, which is what caused real, live 403s and "not
    enabled"/"no permission" errors this session despite the classic table
    itself being unquestionably correct. Writing this directly, right after
    the classic bind, makes access effective immediately instead of waiting
    on that mirror. The `powerpagecomponents` record shares the SAME id as
    its `mspp_entitypermission` row (confirmed live) — best-effort: a site
    without the Enhanced Data Model (no matching component) just has
    nothing to update here, so this silently no-ops rather than failing the
    caller's main flow."""
    try:
        res = dv_request(
            "GET",
            f"{cfg['envUrl']}/api/data/v9.2/powerpagecomponents({permission_id})",
            token,
            params={"$select": "content"},
            timeout=30.0,
        )
        if res.status_code >= 400:
            return
        raw = res.json().get("content")
        content = json.loads(raw) if raw else {}
        content["entityname"] = entity_display_name
        content["entitylogicalname"] = entity_logical_name
        content["scope"] = GLOBAL_SCOPE
        content["create"] = crud.get("create", False)
        content["read"] = crud.get("read", False)
        content["write"] = crud.get("write", False)
        content["append"] = crud.get("append", False)
        content["appendto"] = crud.get("appendto", False)
        content.setdefault("delete", False)
        content["adx_entitypermission_webrole"] = webrole_ids
        dv_request(
            "PATCH",
            f"{cfg['envUrl']}/api/data/v9.2/powerpagecomponents({permission_id})",
            token,
            json={"content": json.dumps(content)},
            timeout=30.0,
        )
    except Exception:
        pass


def ensure_webapi_access(
    cfg: dict,
    token: str,
    entity_logical_name: str,
    entity_display_name: str,
    field_logical_names: list[str],
    *,
    create: bool = True,
    read: bool = True,
    write: bool = True,
    append: bool = False,
    appendto: bool = False,
    extra_webrole_ids: list[str] | None = None,
) -> None:
    """Makes one table fully Web-API-usable: `enabled` + `fields` site
    settings (create-or-update, ADDING `field_logical_names` to whatever
    fields are already allow-listed rather than replacing them — this table
    may not be one this app created, e.g. a Lookup's target table that
    already has its own Web API access configured for other reasons), and a
    Table Permission — Global scope, bound to Administrators (always) plus
    whatever `extra_webrole_ids` names (e.g. Authenticated Users, for a
    template's own main table — a Service Provider submitting a real report
    row needs Create on a role they actually hold, and this site has no
    dedicated Service Provider role). CRUD/Append flags per the kwargs
    (Create/Read/Write/Append/AppendTo for a table this app owns and end
    users submit into; Read-only, no Append, is passed for a Lookup's target
    table, which only ever needs to be displayed).

    Confirmed live: setting a lookup at row-creation time is an
    "association" under the hood, which Dataverse's security model gates on
    Append (source table) + AppendTo (target table) separately from Create —
    a plain `mspp_create: true` isn't enough once any `@odata.bind` is in
    the POST body.

    Upgrades an already-existing permission (delete + recreate — PATCHing
    these records fails with a platform quirk, same as site settings) when
    it's missing something this call asks for; a call that asks for nothing
    new stays a full no-op, so the Lookup-target-table call path (which
    never asks for append/appendto/extra roles) is untouched and still
    idempotent as before."""
    exists_res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{entity_logical_name}')"
        f"?$select=LogicalName",
        token,
        timeout=30.0,
    )
    if exists_res.status_code >= 400:
        raise DataverseError(
            f"No table named '{entity_logical_name}' exists — refusing to create "
            f"Web API access for a table that isn't real."
        )

    existing_fields_raw = _site_setting_value(
        cfg, token, f"Webapi/{entity_logical_name}/fields"
    )
    existing_fields = {
        f.strip() for f in (existing_fields_raw or "").split(",") if f.strip()
    }
    merged_fields = sorted(existing_fields | set(field_logical_names))

    _upsert_site_setting(cfg, token, f"Webapi/{entity_logical_name}/enabled", "true")
    _upsert_site_setting(
        cfg,
        token,
        f"Webapi/{entity_logical_name}/fields",
        ",".join(merged_fields),
    )

    perm_url = (
        f"{cfg['envUrl']}/api/data/v9.2/mspp_entitypermissions"
        f"?$filter=mspp_entitylogicalname eq '{entity_logical_name}' "
        f"and _mspp_websiteid_value eq {cfg['websiteId']}"
        f"&$select=mspp_entitypermissionid,mspp_create,mspp_read,mspp_write,mspp_append,mspp_appendto"
    )
    res = dv_request("GET", perm_url, token, timeout=30.0)
    existing = (res.json().get("value") or [None])[0] if res.status_code < 400 else None

    needs_upgrade = bool(existing) and (
        (append and not existing.get("mspp_append"))
        or (appendto and not existing.get("mspp_appendto"))
        or bool(extra_webrole_ids)
    )
    if existing and not needs_upgrade:
        # The classic table itself needs no change, but its
        # `powerpagecomponents` mirror might still be missing the web role
        # array regardless (confirmed live — this is exactly what silently
        # broke `simform_testtest` submissions and the `contact` admin
        # lookup this session, on a table whose classic CRUD flags were
        # already completely correct). Always re-sync it from the classic
        # table's own current, authoritative state.
        _sync_powerpagecomponent_webroles(
            cfg,
            token,
            existing["mspp_entitypermissionid"],
            _current_webrole_ids_for_permission(
                cfg, token, existing["mspp_entitypermissionid"]
            ),
            entity_logical_name,
            entity_display_name,
            {
                "create": existing.get("mspp_create"),
                "read": existing.get("mspp_read"),
                "write": existing.get("mspp_write"),
                "append": existing.get("mspp_append"),
                "appendto": existing.get("mspp_appendto"),
            },
        )
        return

    merged_create = create or bool(existing and existing.get("mspp_create"))
    merged_read = read or bool(existing and existing.get("mspp_read"))
    merged_write = write or bool(existing and existing.get("mspp_write"))
    merged_append = append or bool(existing and existing.get("mspp_append"))
    merged_appendto = appendto or bool(existing and existing.get("mspp_appendto"))

    if existing:
        dv_request(
            "DELETE",
            f"{cfg['envUrl']}/api/data/v9.2/mspp_entitypermissions"
            f"({existing['mspp_entitypermissionid']})",
            token,
            timeout=30.0,
        )

    body = {
        "mspp_entityname": entity_display_name,
        "mspp_entitylogicalname": entity_logical_name,
        "mspp_scope": GLOBAL_SCOPE,
        "mspp_create": merged_create,
        "mspp_read": merged_read,
        "mspp_write": merged_write,
        "mspp_delete": False,
        "mspp_append": merged_append,
        "mspp_appendto": merged_appendto,
        "mspp_websiteid@odata.bind": f"/mspp_websites({cfg['websiteId']})",
    }
    create_res = dv_request(
        "POST",
        f"{cfg['envUrl']}/api/data/v9.2/mspp_entitypermissions",
        token,
        json=body,
        timeout=30.0,
    )
    if create_res.status_code >= 400:
        raise DataverseError(
            f"Table permission creation failed for {entity_logical_name} "
            f"({create_res.status_code})"
        )
    entity_id_header = create_res.headers.get("odata-entityid", "")
    match = re.search(r"\(([0-9a-fA-F-]{36})\)", entity_id_header)
    perm_id = match.group(1) if match else None
    if not perm_id:
        return

    bound_role_ids = [
        role_id
        for role_id in [cfg["adminWebRoleId"], *(extra_webrole_ids or [])]
        if role_id
    ]
    for role_id in bound_role_ids:
        dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/mspp_entitypermissions({perm_id})"
            f"/mspp_entitypermission_webrole/$ref",
            token,
            json={
                "@odata.id": (
                    f"{cfg['envUrl']}/api/data/v9.2/mspp_webroles({role_id})"
                )
            },
            timeout=30.0,
        )

    # See _sync_powerpagecomponent_webroles' own docs — required on an
    # Enhanced Data Model site (confirmed: this one is) for the binding
    # above to actually take effect, rather than waiting on an unreliable
    # mirror from the classic table.
    _sync_powerpagecomponent_webroles(
        cfg,
        token,
        perm_id,
        bound_role_ids,
        entity_logical_name,
        entity_display_name,
        {
            "create": merged_create,
            "read": merged_read,
            "write": merged_write,
            "append": merged_append,
            "appendto": merged_appendto,
        },
    )


# Dataverse's AttributeTypeName.Value -> the @odata.type needed to PATCH
# that attribute. Same vocabulary build_field_attribute already writes when
# creating each type — this is just its inverse.
_ODATA_TYPE_BY_ATTRIBUTE_TYPE_NAME = {
    "StringType": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
    "MemoType": "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
    "IntegerType": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
    "DecimalType": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
    "MoneyType": "Microsoft.Dynamics.CRM.MoneyAttributeMetadata",
    "BooleanType": "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
    "DateTimeType": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
    "PicklistType": "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
    "MultiSelectPicklistType": "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata",
    "FileType": "Microsoft.Dynamics.CRM.FileAttributeMetadata",
    "ImageType": "Microsoft.Dynamics.CRM.ImageAttributeMetadata",
    "LookupType": "Microsoft.Dynamics.CRM.LookupAttributeMetadata",
}


def _resolve_attribute_info(
    cfg: dict, token: str, entity_logical_name: str, attribute_logical_name: str
) -> dict:
    """Looks up an existing attribute's real type + MetadataId directly from
    Dataverse (rather than trusting whatever the SPA thinks it is) — the
    source of truth for which @odata.type an update PUT needs, including
    correctly telling a single-select Picklist apart from a
    MultiSelectPicklist. `MetadataId` is required in the PUT body for at
    least Picklist attributes (confirmed live: a bare partial body without it
    fails with "A null value was provided for attribute logicalname" — not
    documented, discovered by testing), and is harmless to include for every
    other type, so update_attribute always sends it now."""
    res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{entity_logical_name}')"
        f"/Attributes(LogicalName='{attribute_logical_name}')?$select=AttributeTypeName,MetadataId",
        token,
        timeout=30.0,
    )
    if res.status_code >= 400:
        raise DataverseError(
            f"Could not resolve column type for '{attribute_logical_name}' "
            f"({res.status_code})"
        )
    data = res.json()
    type_name = (data.get("AttributeTypeName") or {}).get("Value")
    odata_type = _ODATA_TYPE_BY_ATTRIBUTE_TYPE_NAME.get(type_name)
    if not odata_type:
        raise DataverseError(
            f"Unsupported column type '{type_name}' for '{attribute_logical_name}' "
            f"— can't update it."
        )
    return {
        "odataType": odata_type,
        "typeName": type_name,
        "metadataId": data.get("MetadataId"),
    }


def update_attribute(
    cfg: dict,
    token: str,
    entity_logical_name: str,
    attribute_logical_name: str,
    display_name: str | None = None,
    description: str | None = None,
    required: bool | None = None,
    min_value: float | None = None,
    max_value: float | None = None,
    precision: int | None = None,
    max_size_kb: int | None = None,
    default_value: bool | None = None,  # Boolean only
    default_option_value: str | None = None,  # Choice's DefaultFormValue
) -> None:
    """Partially updates an EXISTING column. Confirmed live, property by
    property, which of these Dataverse actually allows changing after
    creation (see the plan doc / PR notes for the test results) — every
    param here is one that's safe: DisplayName/Description/RequiredLevel
    (base AttributeMetadata, shared by every type), Min/Max/Precision
    (Integer/Decimal/Money), MaxSizeInKB (File/Image), DefaultValue
    (Boolean), DefaultFormValue (Picklist). Deliberately does NOT support
    AllowMultiple (structurally a different attribute type — PUT rejected
    outright), Boolean TrueOption/FalseOption labels (PUT returns success but
    silently doesn't apply — confirmed via a publish + re-read, worse than an
    error), or Lookup target (relationship retargeting isn't supported at
    all) — callers must never pass those.
    `MSCRM.MergeLabels` keeps this from wiping other languages' labels — it
    only replaces the current language's Label, not the whole LocalizedLabels
    collection. A no-op (no request sent) when nothing was actually passed."""
    if (
        display_name is None
        and description is None
        and required is None
        and min_value is None
        and max_value is None
        and precision is None
        and max_size_kb is None
        and default_value is None
        and default_option_value in (None, "")
    ):
        return

    # Confirmed live: attribute metadata rejects PATCH outright ("does not
    # support http method 'PATCH'") even with the concrete derived type as a
    # URL cast segment. Unlike normal data records, updating one of these
    # requires PUT to that cast segment — Dataverse's metadata API still
    # accepts a partial body here (no SchemaName/AttributeType/etc. needed),
    # it just insists on the PUT verb AND (at least for Picklist) MetadataId
    # actually present in the body. Resolved before building the body (not
    # after, like before) because MinValue/MaxValue's own JSON shape depends
    # on it — see below.
    info = _resolve_attribute_info(
        cfg, token, entity_logical_name, attribute_logical_name
    )
    body: dict = {"MetadataId": info["metadataId"]}
    if display_name is not None:
        body["DisplayName"] = label(display_name)
    if description is not None:
        body["Description"] = label(description)
    if required is not None:
        body["RequiredLevel"] = required_level(
            "ApplicationRequired" if required else "None"
        )
    if min_value is not None:
        # Confirmed live: an Integer attribute's MinValue/MaxValue are
        # Edm.Int32 and Dataverse's OData layer rejects a decimal literal
        # for them outright ("Cannot convert the literal '0.0' to the
        # expected type 'Edm.Int32'") — this call's own params are typed
        # `float` throughout (correct for Decimal/Currency, whose
        # MinValue/MaxValue genuinely are Edm.Decimal), so an Integer field
        # needs its value cast back down to a plain int here, per-attribute,
        # using the real type just resolved above.
        body["MinValue"] = (
            int(min_value) if info["typeName"] == "IntegerType" else min_value
        )
    if max_value is not None:
        body["MaxValue"] = (
            int(max_value) if info["typeName"] == "IntegerType" else max_value
        )
    if precision is not None:
        body["Precision"] = precision
    if max_size_kb is not None:
        body["MaxSizeInKB"] = max_size_kb
    if default_value is not None:
        body["DefaultValue"] = default_value
    if default_option_value not in (None, ""):
        # Resolved through the same numbering used to build the option set
        # at creation time, so it always matches one of the real Values.
        body["DefaultFormValue"] = _picklist_option_value(default_option_value, 0)
    res = dv_request(
        "PUT",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{entity_logical_name}')"
        f"/Attributes(LogicalName='{attribute_logical_name}')/{info['odataType']}",
        token,
        json=body,
        headers={"MSCRM.MergeLabels": "true"},
        timeout=30.0,
    )
    if res.status_code >= 400:
        err_json = res.json() if res.content else {}
        message = (
            (err_json.get("error") or {}).get("message")
            or f"Column update failed ({res.status_code})"
        )
        raise DataverseError(message)


def get_attribute_detail(
    cfg: dict, token: str, entity_logical_name: str, attribute_logical_name: str
) -> dict:
    """One attribute's full type-specific config — MinValue/MaxValue,
    Precision, MaxSizeInKB, DefaultValue, TrueOption/FalseOption labels,
    DefaultFormValue, AllowMultiple, and (for Choice) every option with its
    real Value — whichever apply to its actual type. This is what lets the
    Edit view show an existing field's current settings instead of guessing;
    none of this lives in simform_reportfields — deliberately creation-time-
    only, per the per-type wiring passes throughout this project."""
    info = _resolve_attribute_info(
        cfg, token, entity_logical_name, attribute_logical_name
    )
    needs_option_set = info["typeName"] in (
        "PicklistType",
        "MultiSelectPicklistType",
        "BooleanType",
    )
    url = (
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{entity_logical_name}')"
        f"/Attributes(LogicalName='{attribute_logical_name}')/{info['odataType']}"
    )
    if needs_option_set:
        url += "?$expand=OptionSet"
    res = dv_request("GET", url, token, timeout=30.0)
    if res.status_code >= 400:
        raise DataverseError(
            f"Could not read details for '{attribute_logical_name}' "
            f"({res.status_code})"
        )
    data = res.json()
    detail: dict = {}
    if "MinValue" in data:
        detail["minValue"] = data["MinValue"]
    if "MaxValue" in data:
        detail["maxValue"] = data["MaxValue"]
    if "Precision" in data:
        detail["precision"] = data["Precision"]
    if "MaxSizeInKB" in data:
        detail["maxSizeKb"] = data["MaxSizeInKB"]

    option_set = data.get("OptionSet") or {}
    if info["typeName"] == "BooleanType":
        detail["defaultValue"] = data.get("DefaultValue")
        detail["trueLabel"] = (
            ((option_set.get("TrueOption") or {}).get("Label") or {})
            .get("UserLocalizedLabel", {})
            .get("Label")
        )
        detail["falseLabel"] = (
            ((option_set.get("FalseOption") or {}).get("Label") or {})
            .get("UserLocalizedLabel", {})
            .get("Label")
        )
    if info["typeName"] in ("PicklistType", "MultiSelectPicklistType"):
        detail["allowMultiple"] = info["typeName"] == "MultiSelectPicklistType"
        default_form_value = data.get("DefaultFormValue")
        if default_form_value not in (None, -1):
            detail["defaultOptionValue"] = str(default_form_value)
        detail["options"] = [
            {
                "value": str(o.get("Value")),
                "label": ((o.get("Label") or {}).get("UserLocalizedLabel") or {}).get(
                    "Label"
                ),
            }
            for o in option_set.get("Options") or []
        ]
    return detail


def manage_choice_options(
    cfg: dict,
    token: str,
    entity_logical_name: str,
    attribute_logical_name: str,
    options: list[dict],
) -> None:
    """Applies the desired option list to an EXISTING Choice column — add-
    and-rename only, no delete (removing a live option risks orphaning
    existing data, and wasn't asked for). Each entry in `options` is
    `{value, label}`: a `value` matching one of the column's current real
    Values renames that option via UpdateOptionValue (only when the label
    actually changed — skipped otherwise, and — confirmed live, deliberately
    not attempted — Dataverse has no way to change an existing option's
    Value once created, `UpdateOptionValue` rejects any parameter for it
    outright); anything else (blank, or a value that doesn't match anything
    current — e.g. a row just added on the canvas) adds a new option via
    InsertOptionValue, using that same `value` as the new option's real
    Value (parsed as an int; auto-numbered same as table creation if it
    isn't one)."""
    current = get_attribute_detail(
        cfg, token, entity_logical_name, attribute_logical_name
    )
    current_by_value = {o["value"]: o["label"] for o in current.get("options", [])}

    for i, opt in enumerate(options):
        raw_value = str(opt.get("value") or "").strip()
        new_label = str(opt.get("label") or "").strip()
        if not new_label:
            continue
        if raw_value and raw_value in current_by_value:
            if current_by_value[raw_value] == new_label:
                continue  # unchanged, nothing to do
            res = dv_request(
                "POST",
                f"{cfg['envUrl']}/api/data/v9.2/UpdateOptionValue",
                token,
                json={
                    "AttributeLogicalName": attribute_logical_name,
                    "EntityLogicalName": entity_logical_name,
                    "Value": int(raw_value),
                    "Label": label(new_label),
                    "MergeLabels": False,
                },
                timeout=60.0,
            )
            if res.status_code >= 400:
                err_json = res.json() if res.content else {}
                message = (
                    (err_json.get("error") or {}).get("message")
                    or f"Renaming option '{new_label}' failed ({res.status_code})"
                )
                raise DataverseError(message)
        else:
            # Confirmed live: this call was previously missing `Value`
            # entirely — Dataverse silently auto-numbered the new option
            # using its own internal sequence instead of whatever the admin
            # actually typed into the Value column. Resolve it the same way
            # a brand-new field's initial options are numbered at creation
            # time, so what the admin typed here is what actually lands.
            res = dv_request(
                "POST",
                f"{cfg['envUrl']}/api/data/v9.2/InsertOptionValue",
                token,
                json={
                    "AttributeLogicalName": attribute_logical_name,
                    "EntityLogicalName": entity_logical_name,
                    "Label": label(new_label),
                    "Value": _picklist_option_value(raw_value, i),
                },
                timeout=60.0,
            )
            if res.status_code >= 400:
                err_json = res.json() if res.content else {}
                message = (
                    (err_json.get("error") or {}).get("message")
                    or f"Adding option '{new_label}' failed ({res.status_code})"
                )
                raise DataverseError(message)


def _entity_summary(cfg: dict, token: str, entity_logical_name: str) -> dict | None:
    """A table's display name + primary name column's logical name — used
    to work out which fields a Lookup's target table needs Web API access
    to (at minimum, its primary name column, so the referenced row can be
    identified/displayed at all). Returns None if the table can't be found."""
    res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{entity_logical_name}')"
        f"?$select=DisplayName,PrimaryNameAttribute",
        token,
        timeout=30.0,
    )
    if res.status_code >= 400:
        return None
    data = res.json()
    display_name = ((data.get("DisplayName") or {}).get("UserLocalizedLabel") or {}).get(
        "Label"
    )
    return {
        "displayName": display_name or entity_logical_name,
        "primaryNameAttribute": data.get("PrimaryNameAttribute"),
    }


def _create_lookup_relationships(
    cfg: dict,
    token: str,
    table_schema_name: str,
    table_logical_name: str,
    lookup_fields: list[dict],
) -> None:
    """Creates one relationship (+ embedded Lookup column) per Lookup field
    in `lookup_fields`, referencing `table_logical_name`, then publishes once
    so the new columns are immediately queryable. Shared by
    create_table_with_fields (the referencing entity is the table just
    created) and add_fields_to_table (it already existed). No-op when
    `lookup_fields` is empty."""
    if not lookup_fields:
        return
    for lf in lookup_fields:
        field_schema_name = lf["_schemaName"]
        relationship_schema_name = f"{table_schema_name}_{field_schema_name}"
        create_lookup_relationship(
            cfg,
            token,
            referencing_entity=table_logical_name,
            referenced_entity=lf["_target"],
            relationship_schema_name=relationship_schema_name,
            lookup_schema_name=field_schema_name,
            display_name=lf["displayName"],
            description=lf.get("description"),
        )
    pub_res = dv_request(
        "POST",
        f"{cfg['envUrl']}/api/data/v9.2/PublishAllXml",
        token,
        json={},
        timeout=PUBLISH_TIMEOUT,
    )
    if pub_res.status_code >= 400:
        err_json = pub_res.json() if pub_res.content else {}
        message = (err_json.get("error") or {}).get("message") or "Publish failed"
        raise DataverseError(
            f"Lookup relationship(s) created but publish failed: {message}"
        )


def _grant_lookup_target_access(
    cfg: dict, token: str, lookup_fields: list[dict]
) -> list[str]:
    """Each Lookup's target table also needs Web API access — the SPA can't
    read the referenced record's display value otherwise. A target may be a
    table this app never touched before (e.g. a standard table, or another
    app's custom table), so this is Read-only and additive (see
    ensure_webapi_access's field merging) — it never removes or narrows
    whatever access that table already has, and never grants Create/Write it
    didn't already have. Also binds Authenticated Users (in addition to
    Administrators, always bound) — a Service Provider filling out a real
    report needs to read the target table's own rows to populate the
    Lookup's dropdown, same reasoning as _ensure_activity_report_lookup's
    own extra_webrole_ids. Returns per-target warning strings for whichever
    grants failed (non-fatal — the relationship/column already exist either
    way); shared by create_table_with_fields and add_fields_to_table."""
    warnings: list[str] = []
    for lf in lookup_fields:
        target = lf["_target"]
        summary = _entity_summary(cfg, token, target)
        needed_fields = set()
        if summary and summary.get("primaryNameAttribute"):
            needed_fields.add(summary["primaryNameAttribute"])
        display_field = (lf.get("lookupDisplayField") or "").strip().lower()
        if display_field:
            needed_fields.add(display_field)
        # The target's own id — Dataverse's Web API only ever returns
        # $select-ed columns, and the dropdown's options need the real id
        # to bind against just as much as the label text.
        needed_fields.add(f"{target}id")
        try:
            ensure_webapi_access(
                cfg,
                token,
                target,
                (summary or {}).get("displayName") or target,
                sorted(needed_fields),
                create=False,
                read=True,
                write=False,
                # Confirmed live the hard way: binding a Lookup at
                # row-creation is an "association" under Dataverse's
                # security model, gated on Append (source table — already
                # covered by the main table's own ensure_webapi_access
                # call) AND AppendTo (TARGET table — this one). Missing
                # here entirely until now, this function only ever granted
                # Read on the target, so every dynamic template's Lookup
                # field failed to submit with "You don't have permission
                # to associate or disassociate table X to Y" the moment
                # someone actually tried filling it out — the relationship
                # and columns looked completely fine, only the real
                # row-creation write ever surfaced it.
                appendto=True,
                extra_webrole_ids=[
                    _lookup_webrole_id(cfg, token, "Authenticated Users")
                ],
            )
        except Exception as target_access_err:
            warnings.append(f"{target}: {target_access_err}")
    return warnings


def create_lookup_relationship(
    cfg: dict,
    token: str,
    referencing_entity: str,
    referenced_entity: str,
    relationship_schema_name: str,
    lookup_schema_name: str,
    display_name: str,
    description: str | None,
) -> None:
    """Creates a real N:1 relationship from `referencing_entity` (this new
    table — the "many" side) to `referenced_entity` (the admin's chosen
    target table — the "one" side), with an embedded Lookup attribute. This
    is the only way a real Dataverse Lookup column comes into existence — it
    can't be created as a plain attribute the way every other field type
    can (see build_field_attribute's Choice/Boolean/etc branches), and both
    entities must already exist before this call. Requires a follow-up
    PublishAllXml before the new column is queryable via the Web API."""
    body = {
        "SchemaName": relationship_schema_name,
        "@odata.type": "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata",
        "ReferencedEntity": referenced_entity,
        "ReferencingEntity": referencing_entity,
        # Deliberately inert on delete/share/assign — a POC-built lookup
        # shouldn't cascade deletes or reparent records just because an
        # admin pointed a field at another table.
        "CascadeConfiguration": {
            "Assign": "NoCascade",
            "Delete": "RemoveLink",
            "Merge": "NoCascade",
            "Reparent": "NoCascade",
            "Share": "NoCascade",
            "Unshare": "NoCascade",
        },
        "Lookup": {
            "@odata.type": "Microsoft.Dynamics.CRM.LookupAttributeMetadata",
            "AttributeType": "Lookup",
            "AttributeTypeName": {"Value": "LookupType"},
            "DisplayName": label(display_name),
            "Description": label(description or display_name),
            "SchemaName": lookup_schema_name,
        },
    }
    res = dv_request(
        "POST",
        f"{cfg['envUrl']}/api/data/v9.2/RelationshipDefinitions",
        token,
        json=body,
        timeout=PUBLISH_TIMEOUT,
    )
    if res.status_code >= 400:
        err_json = res.json() if res.content else {}
        message = (
            (err_json.get("error") or {}).get("message")
            or f"Lookup relationship creation failed ({res.status_code})"
        )
        raise DataverseError(message)


def _grid_lookup_schema_name(template_table_logical_name: str) -> str:
    """Deterministic Lookup SchemaName for the real column a grid table
    (e.g. simform_materialloaddetails) uses to link a row back to one
    specific template's own table — strip the "simform_" prefix, PascalCase
    what's left (e.g. simform_testtest -> simform_Testtest).

    grid-submissions.ts (the SPA) computes this SAME name independently
    (deriveGridParentLookupNavProperty) to know which navigation property to
    bind through when it writes a grid's real rows — the two are kept in
    sync by both following this exact rule, not by one asking the other.
    Change one, change both."""
    bare = template_table_logical_name
    if bare.startswith("simform_"):
        bare = bare[len("simform_") :]
    return "simform_" + bare[:1].upper() + bare[1:]


def ensure_grid_table_lookup(
    cfg: dict, token: str, grid_table_logical_name: str, template_table_logical_name: str
) -> str:
    """Adds a real Lookup column on `grid_table_logical_name` pointing back
    at `template_table_logical_name` — the parent-link a Grid Table field's
    real rows bind through when the SPA writes them (see
    src/data/grid-submissions.ts's writeGridRows). A real Dataverse Lookup
    can only ever target ONE table, so a grid table reused across several
    templates (e.g. simform_materialloaddetails, picked by both
    Transportation Reimbursement and any admin-built template) needs its own
    such column per template — this is what provisions each new one, called
    once per (grid table, template table) pair the moment an admin picks
    that grid table for a Grid Table field and saves.

    Idempotent: looks for an EXISTING relationship between these two exact
    tables first (by inspecting the grid table's real relationships,
    matched on ReferencedEntity — not by assuming the derived name, so a
    hand-named exception like simform_TransportationReport is still found
    and reused rather than duplicated) before creating a new one.

    Grants Web API access on BOTH sides with Append AND AppendTo — confirmed
    live the hard way debugging simform_gridconfiguration ->
    simform_reportfields: unlike a normal Lookup (source gets
    Append+AppendTo, target gets AppendTo-only — see
    _ensure_activity_report_lookup + _grant_lookup_target_access), THIS
    relationship needed AppendTo on both the grid table AND the template
    table before the association write stopped 403ing, so both get the full
    set here rather than trying to split it narrower.

    Raises on failure, unlike _ensure_activity_report_lookup — this isn't a
    nice-to-have a template can silently do without; it's the one thing that
    makes a Grid Table field's data save for real, so a failure here should
    reach the admin, not vanish into a dead-end grid nobody notices doesn't
    work. Returns the Lookup's real SchemaName (whichever ends up correct —
    freshly created, or already there)."""
    rel_res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{grid_table_logical_name}')"
        f"/ManyToOneRelationships?$select=SchemaName,ReferencingAttribute,ReferencedEntity"
        f"&$filter=ReferencedEntity eq '{template_table_logical_name}'",
        token,
        timeout=30.0,
    )
    if rel_res.status_code >= 400:
        raise DataverseError(
            f"Could not inspect {grid_table_logical_name}'s relationships "
            f"({rel_res.status_code})"
        )
    existing = (rel_res.json().get("value") or [None])[0]

    if existing:
        # ReferencingAttribute here is the all-lowercase LogicalName, not
        # the schema-cased form @odata.bind needs — fetch the attribute's
        # own metadata for its real SchemaName instead of assuming it
        # matches _grid_lookup_schema_name's derivation (it won't, for a
        # relationship hand-named before this convention existed, like
        # simform_TransportationReport).
        attr_res = dv_request(
            "GET",
            f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{grid_table_logical_name}')"
            f"/Attributes(LogicalName='{existing['ReferencingAttribute']}')?$select=SchemaName",
            token,
            timeout=30.0,
        )
        if attr_res.status_code >= 400:
            raise DataverseError(
                f"Could not resolve real casing for {existing['ReferencingAttribute']} "
                f"({attr_res.status_code})"
            )
        lookup_schema_name = attr_res.json()["SchemaName"]
    else:
        lookup_schema_name = _grid_lookup_schema_name(template_table_logical_name)
        create_lookup_relationship(
            cfg,
            token,
            referencing_entity=grid_table_logical_name,
            referenced_entity=template_table_logical_name,
            relationship_schema_name=(
                f"{grid_table_logical_name}_{lookup_schema_name}_"
                f"{template_table_logical_name}"
            ),
            lookup_schema_name=lookup_schema_name,
            display_name=lookup_schema_name.replace("simform_", ""),
            description=(
                "Internal — links this row back to its parent "
                f"{template_table_logical_name} record. Auto-provisioned the "
                "first time a Grid Table field on that template picked this "
                "grid table."
            ),
        )
        pub_res = dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/PublishAllXml",
            token,
            json={},
            timeout=PUBLISH_TIMEOUT,
        )
        if pub_res.status_code >= 400:
            raise DataverseError(
                f"Grid table lookup created but publish failed ({pub_res.status_code})"
            )

    lower_cased = lookup_schema_name.lower()
    shadow_form = f"_{lower_cased}_value"
    auth_users_id = _lookup_webrole_id(cfg, token, "Authenticated Users")

    ensure_webapi_access(
        cfg,
        token,
        grid_table_logical_name,
        grid_table_logical_name,
        [lower_cased, lookup_schema_name, shadow_form],
        append=True,
        appendto=True,
        extra_webrole_ids=[auth_users_id] if auth_users_id else None,
    )
    ensure_webapi_access(
        cfg,
        token,
        template_table_logical_name,
        template_table_logical_name,
        [],
        append=True,
        appendto=True,
        extra_webrole_ids=[auth_users_id] if auth_users_id else None,
    )

    return lookup_schema_name


def _ensure_activity_report_lookup(
    cfg: dict, token: str, table_logical_name: str, relationship_schema_name: str
) -> None:
    """Adds a hidden Lookup column (simform_ActivityReport) on a template's
    main table, pointing back at simform_serviceprovideractivityreports —
    the real Dataverse relationship a submission's answers row is found
    through. The SPA looks a submission's answers up by *filtering* this
    table on `_simform_activityreport_value eq <the summary row's own id>`
    (see template-submissions.ts's loadSubmittedAnswers) rather than storing
    any pointer on the summary/header side — the header's own id already
    uniquely identifies its one answers row, so no separate "which row"
    column is needed there at all.

    Deliberately never becomes a simform_reportfields catalog row — created
    directly against Dataverse metadata, bypassing the normal field-creation
    path entirely — so it's invisible to the admin Template Builder (no
    canvas entry, never shown in Structure/Preview) and to Service Providers
    filling the form out. Purely internal plumbing this backend manages on
    its own.

    Idempotent (a no-op once the lookup already exists) and best-effort: any
    failure here is swallowed rather than raised, since this runs alongside
    the real table/field creation the caller is actually there for, and a
    template must stay fully usable even if this one internal convenience
    couldn't be added (the "View" answers page just has nothing to show)."""
    try:
        exists_res = dv_request(
            "GET",
            f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')"
            f"/Attributes(LogicalName='{ACTIVITY_REPORT_LOOKUP_LOGICAL_NAME}')?$select=LogicalName",
            token,
            timeout=30.0,
        )
        if exists_res.status_code < 400:
            return  # already there — nothing to do

        create_lookup_relationship(
            cfg,
            token,
            referencing_entity=table_logical_name,
            referenced_entity=ACTIVITY_REPORTS_TABLE,
            relationship_schema_name=relationship_schema_name,
            lookup_schema_name="simform_ActivityReport",
            display_name="Activity Report",
            description=(
                "Internal — links this row back to its Service Provider "
                "Activity Reports summary record. Never shown to admins or "
                "Service Providers."
            ),
        )
        pub_res = dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/PublishAllXml",
            token,
            json={},
            timeout=PUBLISH_TIMEOUT,
        )
        if pub_res.status_code >= 400:
            raise DataverseError("Activity Report lookup created but publish failed")

        # This new column needs the same Create/Append/AppendTo a Service
        # Provider already has on the rest of this table (see the main
        # ensure_webapi_access call this runs alongside) — same reasoning:
        # binding a lookup at row-creation time is an association, gated
        # separately from plain Create. Needs all THREE allow-listed forms,
        # same fix as every other Lookup field in this file (confirmed live
        # the hard way on simform_transportationreimbursementactivityreport,
        # which only had the bare form): the bare lowercase logicalName
        # alone isn't enough — submitTemplateRow's own @odata.bind key is
        # the real case-sensitive schemaName ("simform_ActivityReport"),
        # rejected outright without it allow-listed too, and
        # loadSubmittedAnswers' own $filter on the shadow `_..._value` form
        # 403s without THAT allow-listed. This was previously missing here
        # (only ever patched in live, ad-hoc, for simform_testtest) — now
        # fixed for every table going forward.
        ensure_webapi_access(
            cfg,
            token,
            table_logical_name,
            table_logical_name,
            [
                ACTIVITY_REPORT_LOOKUP_LOGICAL_NAME,
                "simform_ActivityReport",
                f"_{ACTIVITY_REPORT_LOOKUP_LOGICAL_NAME}_value",
            ],
            append=True,
            appendto=True,
            extra_webrole_ids=[
                _lookup_webrole_id(cfg, token, "Authenticated Users")
            ],
        )
    except Exception:
        pass


def _mark_configuration_error(cfg: dict, token: str, config_id: str, message: str) -> None:
    """Best-effort: on any failure after the token was acquired, leave a
    breadcrumb on the configuration row (still Inactive) instead of letting
    it sit forever unexplained — this call runs detached from any HTTP
    response by the time most failures happen. Swallows its own errors."""
    try:
        dv_request(
            "PATCH",
            f"{cfg['envUrl']}/api/data/v9.2/{CONFIG_ENTITY_SET}({config_id})",
            token,
            json={"simform_templatelogicalname": f"ERROR: {message[:200]}"},
            timeout=30.0,
        )
    except Exception:
        pass


def create_table_with_fields(payload: dict) -> dict:
    """Creates one real Dataverse table (with a primary Name column) plus one
    column per field in `payload['fields']` (skipping any `type == 'Lookup'`
    — see build_field_attribute), provisions any Grid Table field's parent
    Lookup (`payload['gridTableFields']`, a list of grid table logical
    names — see ensure_grid_table_lookup), publishes once, then writes the
    resulting logical names back into the two catalog tables:
      - simform_reportconfigurations(configurationId): simform_templatelogicalname
        + simform_displaystatus = Active
      - simform_reportfieldses(fieldRecordId) per non-skipped field:
        simform_fieldlogicalname

    The Grid Table Lookup step happens BEFORE simform_displaystatus is
    written Active, deliberately — confirmed live that leaving it for the
    SPA to do afterward (its own follow-up call, best-effort/fire-and-
    forget) let a template show "Active" up to ~100s before its Grid Table
    field's Lookup actually existed: this function's own Active write goes
    straight to Dataverse, and Power Pages' Web API cache (what the admin's
    list page actually reads) doesn't stay stale nearly long enough to hide
    that gap — so a Service Provider could open a template that LOOKS ready
    and have its Grid Table rows silently fail to save (see
    grid-submissions.ts's writeGridRows). Doing it here instead means
    Active is never written at all until every Grid Table field's Lookup
    genuinely exists.

    On failure, best-effort marks the configuration row with an error
    sentinel (see _mark_configuration_error) rather than leaving it silently
    stuck on Inactive forever, then re-raises.
    """
    cfg = load_config()
    if not cfg:
        raise DataverseError(
            "Backend is missing DATAVERSE_ENV_URL / DATAVERSE_TENANT_ID / "
            "DATAVERSE_CLIENT_ID / DATAVERSE_CLIENT_SECRET environment variables."
        )

    config_id = payload["configurationId"]
    token = get_token(cfg)

    try:
        prefix = (
            re.sub(
                r"[^a-zA-Z0-9]", "", str(payload.get("prefixRaw") or cfg["defaultPrefix"])
            ).lower()
            or cfg["defaultPrefix"]
        )
        table_display_name = payload["tableDisplayName"]
        primary_display_name = payload.get("primaryDisplayName") or "Name"
        # The admin's Help text for the primary field, if they set one —
        # falls back to the old generic description when they didn't.
        primary_description = (
            payload.get("primaryDescription")
            or f"Primary name column for {table_display_name}"
        )
        table_schema_name = f"{prefix}_{pascal_case(table_display_name)}"
        primary_schema_name = f"{prefix}_{pascal_case(primary_display_name)}"

        attributes = [
            {
                "AttributeType": "String",
                "AttributeTypeName": {"Value": "StringType"},
                "Description": label(primary_description),
                "DisplayName": label(primary_display_name),
                "IsPrimaryName": True,
                "RequiredLevel": required_level("None"),
                "SchemaName": primary_schema_name,
                "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
                "FormatName": {"Value": "Text"},
                "MaxLength": 100,
            }
        ]

        # (fieldRecordId, logicalName | None) — None means skipped.
        field_results: list[dict] = []
        # Lookup fields can't become plain attributes on the entity being
        # built below — each needs its own relationship, created only once
        # this table exists for real (see the loop after publish).
        lookup_fields: list[dict] = []
        for f in payload.get("fields", []):
            if f.get("type") == "Lookup":
                target = (f.get("lookupTargetEntity") or "").strip().lower()
                if not target:
                    raise ValueError(
                        f"Lookup field \"{f.get('displayName', '?')}\" has no "
                        f"target entity selected."
                    )
                lookup_fields.append(
                    {
                        **f,
                        "_schemaName": f"{prefix}_{pascal_case(f['displayName'])}",
                        "_target": target,
                    }
                )
                continue
            field_schema_name = f"{prefix}_{pascal_case(f['displayName'])}"
            opts = _field_opts(f)
            attributes.append(
                build_field_attribute(
                    field_schema_name,
                    f["displayName"],
                    f.get("description"),
                    f["type"],
                    opts,
                )
            )
            field_results.append(
                {
                    "fieldRecordId": f["fieldRecordId"],
                    "logicalName": field_schema_name.lower(),
                    "skipped": False,
                    # Written back to simform_reportfields alongside the
                    # logical name below — see simform_allowmultiple's own
                    # column comment for why the catalog needs to carry
                    # this rather than the SP-facing template-catalog.ts
                    # calling back into the (production-unreachable) Python
                    # backend to find out.
                    "allowMultiple": bool(opts.get("allowMultiple")),
                    # Whole Number/Decimal/Currency's real column range —
                    # simform_minvalue/simform_maxvalue mirror it into the
                    # catalog (added specifically so a Service Provider's
                    # browser can validate against it before submit,
                    # without needing the production-unreachable Python
                    # metadata backend — see loadTemplateForFilling and
                    # validateStepFields). None for every other type, which
                    # leaves those two catalog columns unset, same as not
                    # sending them at all.
                    "minValue": opts.get("minValue"),
                    "maxValue": opts.get("maxValue"),
                }
            )

        entity_body = {
            "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata",
            "SchemaName": table_schema_name,
            "Description": label(f"Admin-built report table: {table_display_name}"),
            "DisplayName": label(table_display_name),
            "DisplayCollectionName": label(f"{table_display_name}s"),
            "HasActivities": False,
            "HasNotes": False,
            "IsActivity": False,
            "OwnershipType": payload.get("ownershipType") or cfg["defaultOwnershipType"],
            "Attributes": attributes,
        }

        create_headers = {}
        solution_name = payload.get("solutionUniqueName") or cfg["solutionUniqueName"]
        if solution_name:
            create_headers["MSCRM.SolutionUniqueName"] = solution_name

        create_res = dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions",
            token,
            headers=create_headers,
            json=entity_body,
        )
        if create_res.status_code >= 400:
            err_json = create_res.json() if create_res.content else {}
            message = (
                (err_json.get("error") or {}).get("message")
                or f"Table creation failed ({create_res.status_code})"
            )
            raise DataverseError(message)

        pub_res = dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/PublishAllXml",
            token,
            json={},
            timeout=PUBLISH_TIMEOUT,
        )
        if pub_res.status_code >= 400:
            err_json = pub_res.json() if pub_res.content else {}
            message = (err_json.get("error") or {}).get("message") or "Publish failed"
            raise DataverseError(f"Table created but publish failed: {message}")

        table_logical_name = table_schema_name.lower()

        # Every template's main table gets this hidden Lookup automatically
        # — see _ensure_activity_report_lookup's own docs for why (and why
        # it's never one of the admin's own visible fields).
        _ensure_activity_report_lookup(
            cfg, token, table_logical_name, f"{table_schema_name}_ActivityReport"
        )

        # Now that the referencing table is real, create one relationship
        # (+ its embedded Lookup column) per Lookup field. The target table
        # can be anything in the org (see list_tables()) — it always already
        # exists, since it's picked from a live list, never typed freely.
        _create_lookup_relationships(
            cfg, token, table_schema_name, table_logical_name, lookup_fields
        )
        for lf in lookup_fields:
            field_results.append(
                {
                    "fieldRecordId": lf["fieldRecordId"],
                    "logicalName": lf["_schemaName"].lower(),
                    "skipped": False,
                    # The real, case-sensitive relationship SchemaName —
                    # submitTemplateRow needs this exact casing for
                    # @odata.bind (confirmed live: the lowercase logical
                    # name above is rejected outright for that).
                    "schemaName": lf["_schemaName"],
                }
            )

        # Each Lookup's target table also needs Web API access — see
        # _grant_lookup_target_access's own docstring for why this is
        # Read-only and additive.
        lookup_target_warnings = _grant_lookup_target_access(cfg, token, lookup_fields)

        # A Grid Table field never gets a real column on THIS table (see
        # build_field_attribute/the SPA's own isGridTable doc) — instead it
        # needs a parent Lookup on the GRID table pointing back here, which
        # can only be created now that table_logical_name is real. Done
        # here, before simform_displaystatus is written Active below — see
        # this function's own doc for why that order matters.
        #
        # Retried on the same "EntityCustomization" lock-contention error the
        # SPA's own backendFetch already retries (backend-api.ts's
        # isCustomizationLockError) — confirmed live this is genuinely
        # transient in this org, and NOT swallowing it here matters: an
        # earlier version caught every failure and pressed on to write
        # Active regardless, so one transient lock hit during this step
        # still let a template show "Active" with its Grid Table field
        # permanently unable to save (the table itself is already real by
        # this point, so a plain retry of the whole /api/create-template
        # call from the SPA can't fix it either — creating the table again
        # just 400s on the now-duplicate schema name). Genuinely non-lock
        # errors (a real, non-transient failure) are NOT retried — those go
        # straight to grid_table_warnings below.
        grid_table_warnings: list[str] = []
        for grid_table_logical_name in payload.get("gridTableFields") or []:
            for attempt in range(1, 5):
                try:
                    ensure_grid_table_lookup(
                        cfg, token, grid_table_logical_name, table_logical_name
                    )
                    break
                except Exception as grid_err:
                    is_lock_error = "entitycustomization" in str(grid_err).lower()
                    if is_lock_error and attempt < 4:
                        time.sleep(3 * (2 ** (attempt - 1)))
                        continue
                    grid_table_warnings.append(
                        f"{grid_table_logical_name}: {grid_err}"
                    )
                    break

        # Only mark Active once every Grid Table field's Lookup genuinely
        # exists — see this function's own doc for why. A leftover warning
        # here means at least one didn't; the table/columns/config's real
        # table name still get written (an admin needs that to open the
        # template again and re-save, which retries this same step,
        # ensure_grid_table_lookup being idempotent), but the config stays
        # exactly as inactive as it already was rather than lying about
        # being ready — same "Creating…" the list page already shows.
        config_patch: dict = {"simform_templatelogicalname": table_logical_name}
        if not grid_table_warnings:
            config_patch["simform_displaystatus"] = DISPLAY_STATUS_ACTIVE
        config_res = dv_request(
            "PATCH",
            f"{cfg['envUrl']}/api/data/v9.2/{CONFIG_ENTITY_SET}({config_id})",
            token,
            json=config_patch,
        )
        if config_res.status_code >= 400:
            raise DataverseError(
                f"Table/columns created ({table_logical_name}) but writing back to "
                f"the configuration record failed ({config_res.status_code})."
            )

        for fr in field_results:
            if fr["skipped"]:
                continue
            field_res = dv_request(
                "PATCH",
                f"{cfg['envUrl']}/api/data/v9.2/{FIELDS_ENTITY_SET}({fr['fieldRecordId']})",
                token,
                json={
                    "simform_fieldlogicalname": fr["logicalName"],
                    "simform_allowmultiple": fr.get("allowMultiple", False),
                    "simform_fieldschemaname": fr.get("schemaName"),
                    "simform_minvalue": fr.get("minValue"),
                    "simform_maxvalue": fr.get("maxValue"),
                },
            )
            if field_res.status_code >= 400:
                raise DataverseError(
                    f"Column '{fr['logicalName']}' was created but writing its logical "
                    f"name back to the field record failed ({field_res.status_code})."
                )

        # The new table is real now, but still invisible to the Web API until
        # it has an `enabled`/`fields` site setting and a table permission —
        # see ensure_webapi_access. Failing this step shouldn't undo the
        # table/columns that already exist, so it's reported but non-fatal.
        try:
            all_field_logical_names = [primary_schema_name.lower()] + [
                fr["logicalName"] for fr in field_results if not fr["skipped"]
            ]
            # A Lookup field's own allow-listed name needs THREE forms —
            # see _add_fields_to_table's identical fix for why: the plain
            # lowercase logicalName (still needed for @odata.bind's target
            # in some contexts), the real case-sensitive schemaName
            # (submitTemplateRow's @odata.bind key — Power Pages' allow-
            # list check is case-sensitive, confirmed live the lowercase
            # form alone is rejected for that), and `_<logicalname>_value`
            # (confirmed live: a bare `$select=<logicalname>` on a Lookup
            # column doesn't just 403 — it fails the WHOLE query with a
            # generic "unexpected error" — Power Pages' own allow-list
            # check for $select on a Lookup wants the shadow-value form).
            for fr in field_results:
                if not fr["skipped"] and fr.get("schemaName"):
                    if fr["schemaName"] != fr["logicalName"]:
                        all_field_logical_names.append(fr["schemaName"])
                    all_field_logical_names.append(f'_{fr["logicalName"]}_value')
            # This is a template's own main table — the Service Provider
            # actually filling it out and submitting a real row needs
            # Create (+ Append/AppendTo, since binding a lookup at creation
            # is an association under the hood) for a role they hold; this
            # site has none dedicated, so Authenticated Users stands in.
            ensure_webapi_access(
                cfg,
                token,
                table_logical_name,
                table_display_name,
                all_field_logical_names,
                append=True,
                appendto=True,
                extra_webrole_ids=[
                    _lookup_webrole_id(cfg, token, "Authenticated Users")
                ],
            )
        except Exception as access_err:
            return {
                "success": True,
                "tableLogicalName": table_logical_name,
                "fields": field_results,
                "webApiAccessWarning": str(access_err),
                **(
                    {"lookupTargetAccessWarnings": lookup_target_warnings}
                    if lookup_target_warnings
                    else {}
                ),
                **(
                    {"gridTableWarnings": grid_table_warnings}
                    if grid_table_warnings
                    else {}
                ),
            }

        return {
            "success": True,
            "tableLogicalName": table_logical_name,
            "fields": field_results,
            **(
                {"gridTableWarnings": grid_table_warnings}
                if grid_table_warnings
                else {}
            ),
            **(
                {"lookupTargetAccessWarnings": lookup_target_warnings}
                if lookup_target_warnings
                else {}
            ),
        }
    except Exception as err:
        _mark_configuration_error(cfg, token, config_id, str(err))
        raise


def add_fields_to_table(payload: dict) -> dict:
    """Adds fields to a table that already exists for real — the "Edit
    template" counterpart to create_table_with_fields. Mirrors its async UX:
    the SPA sets simform_displaystatus to Modifying itself before firing
    this (unawaited) and navigating away immediately; this function reverts
    it back to Active once done — on success *or* failure, via `finally`, so
    a template never gets stuck showing "Modifying…" forever. Requires
    `payload['configurationId']` for that revert (a no-op, best-effort, if
    somehow missing — never lets a revert failure mask the real result).

    Unlike create_table_with_fields, a failure here does NOT call
    _mark_configuration_error — that stamps an ERROR sentinel over
    simform_templatelogicalname, which would corrupt an already-correct
    configuration. It just raises; main.py turns that into a 400 — with the
    SPA no longer awaiting this call, that 400 only ever reaches a
    `.catch(console.error)`, the same best-effort philosophy the create flow
    already uses for its own background failures."""
    config_id = payload.get("configurationId")
    cfg = load_config()
    if not cfg:
        raise DataverseError(
            "Backend is missing DATAVERSE_ENV_URL / DATAVERSE_TENANT_ID / "
            "DATAVERSE_CLIENT_ID / DATAVERSE_CLIENT_SECRET environment variables."
        )
    token = get_token(cfg)
    try:
        return _add_fields_to_table(cfg, token, payload)
    finally:
        if config_id:
            try:
                dv_request(
                    "PATCH",
                    f"{cfg['envUrl']}/api/data/v9.2/{CONFIG_ENTITY_SET}({config_id})",
                    token,
                    json={"simform_displaystatus": DISPLAY_STATUS_ACTIVE},
                    timeout=30.0,
                )
            except Exception:
                pass


def _add_fields_to_table(cfg: dict, token: str, payload: dict) -> dict:
    table_logical_name = payload["tableLogicalName"].strip().lower()
    # Recovered from the table's own logical name (prefix_PascalCaseName),
    # not re-read from .env — the table may have been created under a
    # different DEFAULT_TABLE_PREFIX than whatever's configured today.
    prefix = table_logical_name.split("_", 1)[0] or cfg["defaultPrefix"]

    # Retrofits the hidden Activity-Report-linking Lookup (see
    # _ensure_activity_report_lookup) onto a template's table the first
    # time it's edited after this existed — idempotent, so this is cheap
    # and safe to call on every edit, not just ones that add new fields.
    _ensure_activity_report_lookup(
        cfg, token, table_logical_name, f"{table_logical_name}_activityreportlink"
    )

    create_headers = {}
    solution_name = payload.get("solutionUniqueName") or cfg["solutionUniqueName"]
    if solution_name:
        create_headers["MSCRM.SolutionUniqueName"] = solution_name

    field_results: list[dict] = []  # only ever populated for isNew fields
    lookup_fields: list[dict] = []
    added_plain_attribute = False
    # One existing field's update_attribute call throwing (a genuinely
    # unsupported change, a transient Dataverse error, etc.) must not sacrifice
    # every OTHER field queued after it in this same save — confirmed exactly
    # this happened live: an Integer field's Min/Max bug aborted the whole
    # loop, silently skipping every field listed after it. Each existing
    # field's update is now isolated; failures collect here instead.
    field_update_warnings: list[str] = []

    for f in payload.get("fields", []):
        if not f.get("isNew"):
            existing_logical_name = (f.get("logicalName") or "").strip().lower()
            if not existing_logical_name:
                raise ValueError(
                    f"Field \"{f.get('displayName', '?')}\" is marked as "
                    f"existing but has no real column logical name to update."
                )
            try:
                update_attribute(
                    cfg,
                    token,
                    table_logical_name,
                    existing_logical_name,
                    display_name=f.get("displayName"),
                    description=f.get("description"),
                    required=f.get("required"),
                    min_value=f.get("minValue"),
                    max_value=f.get("maxValue"),
                    precision=f.get("precision"),
                    max_size_kb=f.get("maxSizeKb"),
                    default_value=f.get("defaultValue"),
                    default_option_value=f.get("defaultOptionValue"),
                )
                # Choice options: add-and-rename only (see
                # manage_choice_options' own docs) — nothing to do when the
                # field has no options list (every non-Choice type, or a
                # Choice field left untouched).
                if f.get("options"):
                    manage_choice_options(
                        cfg,
                        token,
                        table_logical_name,
                        existing_logical_name,
                        f["options"],
                    )
                # Backfills simform_allowmultiple/minvalue/maxvalue for a
                # field saved before those columns existed (or before this
                # app tracked them at all) — read straight off the real
                # attribute rather than trusting anything the SPA sent, so
                # a stale/never-set value in the catalog gets corrected the
                # next time this field's template is opened and saved, same
                # idempotent "cheap to call on every edit" pattern as
                # _ensure_activity_report_lookup. Every type gets this, not
                # just Choice/numeric — resolving False/None for a type
                # that doesn't have that setting is exactly what it should
                # already be. get_attribute_detail also picks up whatever
                # Min/Max update_attribute just applied above (or the
                # column's pre-existing range, if this save didn't touch
                # it), so simform_minvalue/maxvalue always mirror the real
                # column exactly — this is what lets a Service Provider's
                # browser validate against it later without ever reaching
                # the production-unreachable Python metadata backend
                # itself (see loadTemplateForFilling/validateStepFields).
                detail = get_attribute_detail(
                    cfg, token, table_logical_name, existing_logical_name
                )
                catalog_backfill_res = dv_request(
                    "PATCH",
                    f"{cfg['envUrl']}/api/data/v9.2/{FIELDS_ENTITY_SET}({f['fieldRecordId']})",
                    token,
                    json={
                        "simform_allowmultiple": bool(detail.get("allowMultiple")),
                        "simform_minvalue": detail.get("minValue"),
                        "simform_maxvalue": detail.get("maxValue"),
                    },
                    timeout=30.0,
                )
                if catalog_backfill_res.status_code >= 400:
                    raise DataverseError(
                        "Could not backfill simform_allowmultiple/minvalue/maxvalue "
                        f"({catalog_backfill_res.status_code})"
                    )
            except Exception as field_err:
                field_update_warnings.append(
                    f"\"{f.get('displayName') or existing_logical_name}\": {field_err}"
                )
            continue

        if f.get("type") == "Lookup":
            target = (f.get("lookupTargetEntity") or "").strip().lower()
            if not target:
                raise ValueError(
                    f"Lookup field \"{f.get('displayName', '?')}\" has no "
                    f"target entity selected."
                )
            lookup_fields.append(
                {
                    **f,
                    "_schemaName": f"{prefix}_{pascal_case(f['displayName'])}",
                    "_target": target,
                }
            )
            continue

        field_schema_name = f"{prefix}_{pascal_case(f['displayName'])}"
        opts = _field_opts(f)
        attribute = build_field_attribute(
            field_schema_name,
            f["displayName"],
            f.get("description"),
            f["type"],
            opts,
        )
        add_res = dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')"
            f"/Attributes",
            token,
            headers=create_headers,
            json=attribute,
        )
        if add_res.status_code >= 400:
            err_json = add_res.json() if add_res.content else {}
            message = (
                (err_json.get("error") or {}).get("message")
                or f"Adding column '{field_schema_name}' failed ({add_res.status_code})"
            )
            raise DataverseError(message)
        added_plain_attribute = True
        field_results.append(
            {
                "fieldRecordId": f["fieldRecordId"],
                "logicalName": field_schema_name.lower(),
                "skipped": False,
                "allowMultiple": bool(opts.get("allowMultiple")),
                # See create_table_with_fields' identical field for why.
                "minValue": opts.get("minValue"),
                "maxValue": opts.get("maxValue"),
            }
        )

    if added_plain_attribute:
        pub_res = dv_request(
            "POST",
            f"{cfg['envUrl']}/api/data/v9.2/PublishAllXml",
            token,
            json={},
            timeout=PUBLISH_TIMEOUT,
        )
        if pub_res.status_code >= 400:
            err_json = pub_res.json() if pub_res.content else {}
            message = (err_json.get("error") or {}).get("message") or "Publish failed"
            raise DataverseError(f"Column(s) added but publish failed: {message}")

    _create_lookup_relationships(
        cfg, token, table_logical_name, table_logical_name, lookup_fields
    )
    for lf in lookup_fields:
        field_results.append(
            {
                "fieldRecordId": lf["fieldRecordId"],
                "logicalName": lf["_schemaName"].lower(),
                "skipped": False,
                "schemaName": lf["_schemaName"],
            }
        )

    lookup_target_warnings = _grant_lookup_target_access(cfg, token, lookup_fields)

    if field_results:
        table_summary = _entity_summary(cfg, token, table_logical_name)
        try:
            # Same reasoning as create_table_with_fields' own call — this is
            # the template's main table, and a Service Provider submitting a
            # real row into it needs Create/Append/AppendTo for a role they
            # actually hold. Also retroactively upgrades a table created
            # before this was added, the next time a field is added to it.
            # A Lookup field's own allow-listed name needs THREE forms —
            # its all-lowercase logicalName, its real case-sensitive
            # schemaName for the @odata.bind key submitTemplateRow writes
            # through (Power Pages' allow-list check is case-sensitive —
            # confirmed live, a submission's Lookup answer silently failed
            # to save without this), and `_<logicalname>_value` — confirmed
            # live, a bare `$select=<logicalname>` on a Lookup column fails
            # the WHOLE query with a generic "unexpected error" rather than
            # a clean 403, because Power Pages' $select allow-list check
            # for a Lookup wants the shadow-value form instead.
            allow_listed_fields = [fr["logicalName"] for fr in field_results]
            for fr in field_results:
                if fr.get("schemaName"):
                    if fr["schemaName"] != fr["logicalName"]:
                        allow_listed_fields.append(fr["schemaName"])
                    allow_listed_fields.append(f'_{fr["logicalName"]}_value')
            ensure_webapi_access(
                cfg,
                token,
                table_logical_name,
                (table_summary or {}).get("displayName") or table_logical_name,
                allow_listed_fields,
                append=True,
                appendto=True,
                extra_webrole_ids=[
                    _lookup_webrole_id(cfg, token, "Authenticated Users")
                ],
            )
        except Exception as access_err:
            return {
                "success": True,
                "tableLogicalName": table_logical_name,
                "fields": field_results,
                "webApiAccessWarning": str(access_err),
                **(
                    {"lookupTargetAccessWarnings": lookup_target_warnings}
                    if lookup_target_warnings
                    else {}
                ),
                **(
                    {"fieldUpdateWarnings": field_update_warnings}
                    if field_update_warnings
                    else {}
                ),
            }

    for fr in field_results:
        field_res = dv_request(
            "PATCH",
            f"{cfg['envUrl']}/api/data/v9.2/{FIELDS_ENTITY_SET}({fr['fieldRecordId']})",
            token,
            json={
                "simform_fieldlogicalname": fr["logicalName"],
                "simform_allowmultiple": fr.get("allowMultiple", False),
                "simform_fieldschemaname": fr.get("schemaName"),
                "simform_minvalue": fr.get("minValue"),
                "simform_maxvalue": fr.get("maxValue"),
            },
        )
        if field_res.status_code >= 400:
            raise DataverseError(
                f"Column '{fr['logicalName']}' was created but writing its logical "
                f"name back to the field record failed ({field_res.status_code})."
            )

    return {
        "success": True,
        "tableLogicalName": table_logical_name,
        "fields": field_results,
        **(
            {"lookupTargetAccessWarnings": lookup_target_warnings}
            if lookup_target_warnings
            else {}
        ),
        **(
            {"fieldUpdateWarnings": field_update_warnings}
            if field_update_warnings
            else {}
        ),
    }


def list_tables() -> list[dict]:
    """Every real table in the environment — the candidate list for the
    Lookup field type's "Target Entity" picker. Was previously narrowed to
    just this app's own `simform_` tables; now shows everything (including
    standard tables like account/contact) so an admin can point a Lookup at
    any table in the org. Still excludes intersect tables (many-to-many
    junction entities) and logical entities (computed, no real rows) —
    neither can ever be a valid Lookup relationship target.

    Includes each table's `primaryNameAttribute` — the SPA uses this to set
    a Lookup's Display Field automatically (always the target's primary
    name column) instead of making the admin pick one by hand."""
    cfg = load_config()
    if not cfg:
        raise DataverseError("Backend is missing Dataverse environment variables.")
    token = get_token(cfg)
    res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions"
        f"?$select=LogicalName,DisplayName,IsIntersect,IsLogicalEntity,PrimaryNameAttribute",
        token,
        timeout=30.0,
    )
    if res.status_code >= 400:
        raise DataverseError(f"Could not list tables ({res.status_code})")
    tables = []
    for e in res.json().get("value", []):
        if e.get("IsIntersect") or e.get("IsLogicalEntity"):
            continue
        logical_name = e.get("LogicalName") or ""
        display_name = ((e.get("DisplayName") or {}).get("UserLocalizedLabel") or {}).get("Label")
        tables.append(
            {
                "logicalName": logical_name,
                "displayName": display_name or logical_name,
                "primaryNameAttribute": e.get("PrimaryNameAttribute"),
            }
        )
    tables.sort(key=lambda t: t["displayName"].lower())
    return tables


def list_table_fields(logical_name: str) -> list[dict]:
    """One table's columns — the candidate list for the Lookup field type's
    "Display Field" picker, once a target table has been chosen. Also used
    when loading an existing template for editing: `description` is a
    column's real "Help text" — since simform_reportfields never stores it
    (it only ever becomes the real column's Description at creation time),
    this is the only place it can be read back from."""
    cfg = load_config()
    if not cfg:
        raise DataverseError("Backend is missing Dataverse environment variables.")
    token = get_token(cfg)
    res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{logical_name}')"
        f"/Attributes?$select=LogicalName,DisplayName,AttributeType,Description",
        token,
        timeout=30.0,
    )
    if res.status_code >= 400:
        raise DataverseError(f"Could not list fields for {logical_name} ({res.status_code})")
    fields = []
    for a in res.json().get("value", []):
        display_name = ((a.get("DisplayName") or {}).get("UserLocalizedLabel") or {}).get("Label")
        description = ((a.get("Description") or {}).get("UserLocalizedLabel") or {}).get("Label")
        fields.append(
            {
                "logicalName": a.get("LogicalName"),
                "displayName": display_name or a.get("LogicalName"),
                "type": a.get("AttributeType"),
                "description": description,
            }
        )
    fields.sort(key=lambda f: (f["displayName"] or "").lower())
    return fields


def lookup_table(logical_name: str) -> dict:
    cfg = load_config()
    if not cfg:
        raise DataverseError("Backend is missing Dataverse environment variables.")
    token = get_token(cfg)
    res = dv_request(
        "GET",
        f"{cfg['envUrl']}/api/data/v9.2/EntityDefinitions(LogicalName='{logical_name}')"
        f"?$select=SchemaName,LogicalName,DisplayName,ObjectTypeCode,CreatedOn"
        f"&$expand=Attributes($select=LogicalName,SchemaName,AttributeType,IsPrimaryName,DisplayName)",
        token,
        timeout=30.0,
    )
    if res.status_code >= 400:
        return {"found": False, "status": res.status_code, "body": res.text[:500]}
    data = res.json()
    return {
        "found": True,
        "schemaName": data.get("SchemaName"),
        "logicalName": data.get("LogicalName"),
        "createdOn": data.get("CreatedOn"),
        "attributes": [
            {
                "logicalName": a.get("LogicalName"),
                "schemaName": a.get("SchemaName"),
                "type": a.get("AttributeType"),
                "isPrimaryName": a.get("IsPrimaryName"),
                "displayName": (
                    (a.get("DisplayName") or {}).get("UserLocalizedLabel") or {}
                ).get("Label"),
            }
            for a in data.get("Attributes", [])
        ],
    }

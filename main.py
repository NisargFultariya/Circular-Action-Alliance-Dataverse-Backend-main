"""
Dataverse table/column builder API — FastAPI backend.

Endpoints:
  GET  /api/config              -> is the backend configured? (no secrets returned)
  GET  /api/lookup               -> look up an existing table by logical name
  POST /api/create                -> create a table (+ optional field), then publish
  POST /api/create-template       -> create a table + N fields for an admin-built
                                      template, writing results back into
                                      simform_reportconfigurations/simform_reportfieldses
  GET  /api/tables                -> this app's tables (for the Lookup field picker)
  GET  /api/tables/{name}/fields  -> one table's columns (for the Lookup field picker)
  POST /api/ensure-webapi-access  -> create/update a table's Webapi site settings +
                                      table permission (create-template now calls this
                                      automatically; this is the manual/bootstrap version)
  POST /api/ensure-grid-table-lookup -> provision the real Lookup + Web API access a
                                      Grid Table field's rows bind through, for one
                                      (grid table, template table) pair

Run locally:
    uvicorn main:app --reload --port 8942

Deploy: see README.md for the Render.com step-by-step.
"""

import os
from dotenv import load_dotenv

# Loads a local .env file if present (no-op if it doesn't exist — Render sets
# real environment variables instead, so this line does nothing in production).
load_dotenv()

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from dataverse import (
    DataverseError,
    _lookup_webrole_id,
    add_fields_to_table,
    create_table_and_field,
    create_table_with_fields,
    ensure_grid_table_lookup,
    ensure_webapi_access,
    get_attribute_detail,
    get_token,
    list_table_fields,
    list_tables,
    load_config,
    lookup_table,
)

app = FastAPI(title="Dataverse Table Builder API")

# Which origins may call this API from a browser. Set ALLOWED_ORIGINS as a
# comma-separated env var in production (your Power Pages site's real URL) —
# "*" here is a POC-only default, not something to ship with.
allowed_origins = [
    o.strip()
    for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",")
    if o.strip()
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


@app.middleware("http")
async def allow_private_network_access(request, call_next):
    """Grants Chrome's newer "Private Network Access" permission — without
    this, a public HTTPS caller (the deployed Power Pages site, this
    backend's only real caller) reaching into a private/loopback address
    like localhost gets silently blocked by Chrome itself, entirely
    separate from (and in addition to) the CORSMiddleware allow_origins
    check above: "Access to fetch at 'http://localhost:8942/...' ... has
    been blocked by CORS policy: Permission was denied for this request to
    access the `loopback` address space." Confirmed live — the exact same
    backend code, same .env, same org, works from one machine's Chrome and
    not another's purely because of which Chrome version/rollout stage is
    enforcing this spec; CORSMiddleware's own allow_origins="*" already
    says yes, but Starlette's CORSMiddleware has no built-in support for
    this separate, newer header. Added here (a middleware layered OUTSIDE
    CORSMiddleware, since this function is registered after it) so it also
    reaches CORSMiddleware's own short-circuited OPTIONS preflight
    response, not just real GET/POST responses — Chrome checks this header
    on the preflight, before ever sending the real request.
    Unconditional: a normal same-network caller (a request that was never
    going to be flagged as cross-private-network to begin with) simply
    ignores a header it never needed."""
    response = await call_next(request)
    response.headers["Access-Control-Allow-Private-Network"] = "true"
    return response


@app.get("/api/health")
def health():
    """Render's health check hits this — keep it dependency-free."""
    return {"ok": True}


@app.get("/api/config")
def get_config():
    cfg = load_config()
    return {
        "configured": cfg is not None,
        "envUrl": cfg["envUrl"] if cfg else None,
    }


@app.get("/api/lookup")
def get_lookup(name: str):
    try:
        return lookup_table(name.strip().lower())
    except DataverseError as err:
        return {"found": False, "error": str(err)}


class FieldOptions(BaseModel):
    maxLength: int | None = None
    minValue: float | None = None
    maxValue: float | None = None
    precision: int | None = None
    defaultValue: bool | None = None
    trueLabel: str | None = None
    falseLabel: str | None = None
    dateFormat: str | None = None  # 'DateOnly' | 'DateAndTime'
    options: list[dict] | None = None  # Choice: [{label, value}]
    allowMultiple: bool | None = None
    maxSizeKb: int | None = None


class CreateTablePayload(BaseModel):
    # Unset (None) falls back to DEFAULT_TABLE_PREFIX / DEFAULT_OWNERSHIP_TYPE
    # / SOLUTION_UNIQUE_NAME from .env — see dataverse.py's load_config().
    prefixRaw: str | None = None
    tableDisplayName: str
    tableDescription: str | None = None
    tableCollectionName: str | None = None
    ownershipType: str | None = None  # 'UserOwned' | 'OrganizationOwned'
    primaryDisplayName: str | None = None
    primaryMaxLength: int | None = None
    solutionUniqueName: str | None = None
    publish: bool = True

    addField: bool = False
    fieldDisplayName: str | None = None
    fieldDescription: str | None = None
    # One of: Text, Memo, Integer, Decimal, Currency, Boolean, DateTime,
    # Choice, File, Image — matches the SPA's field-type palette 1:1
    # (see README.md's mapping table).
    fieldType: str | None = None
    fieldOptions: FieldOptions | None = None


@app.post("/api/create")
def create_table(payload: CreateTablePayload):
    try:
        body = payload.model_dump(exclude_none=True)
        if payload.fieldOptions:
            body["fieldOptions"] = payload.fieldOptions.model_dump(exclude_none=True)
        return create_table_and_field(body)
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err), "steps": err.steps}) from err
    except ValueError as err:
        raise HTTPException(status_code=400, detail={"error": str(err), "steps": []}) from err


class TemplateFieldPayload(BaseModel):
    fieldRecordId: str
    displayName: str
    # One of: Text, Memo, Integer, Decimal, Currency, Boolean, DateTime,
    # Choice, File, Image, Lookup — the SPA maps its own field-type palette
    # onto these before calling this endpoint. 'Lookup' is accepted but the
    # backend skips real column creation for it (see build_field_attribute).
    type: str
    # The field's Help text from the builder — becomes the real column's
    # Description. Falls back to displayName if omitted (build_field_attribute's
    # own default), not to a blank description.
    description: str | None = None
    options: list[dict] | None = None  # Choice: [{label, value}]
    dateFormat: str | None = None  # DateTime: 'DateOnly' | 'DateAndTime'
    minValue: float | None = None  # Whole Number / Decimal / Currency: value range
    maxValue: float | None = None
    precision: int | None = None  # Decimal / Currency: number of decimal places
    allowMultiple: bool | None = None  # Choice: multi-select picklist
    defaultOptionValue: str | None = None  # Choice: DefaultFormValue
    trueLabel: str | None = None  # Boolean: TrueOption label
    falseLabel: str | None = None  # Boolean: FalseOption label
    defaultValue: bool | None = None  # Boolean: DefaultValue
    maxSizeKb: int | None = None  # File / Image: MaxSizeInKB
    lookupTargetEntity: str | None = None  # Lookup: real relationship's target table
    lookupDisplayField: str | None = None  # Lookup: target column to grant Web API read access to
    required: bool | None = None  # Wires the "Required" checkbox into RequiredLevel


class CreateTemplatePayload(BaseModel):
    configurationId: str  # simform_reportconfiguration GUID, already created
    # Left unset (None), each of these falls back to its DEFAULT_TABLE_PREFIX /
    # DEFAULT_OWNERSHIP_TYPE / SOLUTION_UNIQUE_NAME .env value — see
    # dataverse.py's load_config(). Set here only to override that default
    # for one specific call.
    prefixRaw: str | None = None
    ownershipType: str | None = None
    solutionUniqueName: str | None = None
    tableDisplayName: str
    primaryDisplayName: str | None = None
    primaryDescription: str | None = None  # the primary field's Help text
    fields: list[TemplateFieldPayload] = []


@app.post("/api/create-template")
def create_template(payload: CreateTemplatePayload):
    """Creates the real Dataverse table + columns for an admin-built
    template, then writes the resulting logical names back into
    simform_reportconfigurations/simform_reportfieldses. Meant to be called
    without the caller waiting on the full result — the SPA fires this and
    immediately navigates away; `/admin/templates` finds out it's done by
    polling simform_displaystatus, not from this response.
    """
    try:
        body = payload.model_dump(exclude_none=True)
        return create_table_with_fields(body)
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err
    except ValueError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err


class UpdateFieldPayload(TemplateFieldPayload):
    # True: this field needs a brand-new real column/relationship, exactly
    # like the create flow. False: this field's real column already exists —
    # displayName/description/required get patched onto it, plus whichever
    # of minValue/maxValue/precision/maxSizeKb/defaultValue/defaultOptionValue
    # actually apply to its real type (see dataverse.py's update_attribute —
    # it silently ignores whatever doesn't apply to that attribute type), and
    # `options` add-and-renames Choice options (manage_choice_options) using
    # each option's real Dataverse `value` to know which one to rename.
    # allowMultiple/trueLabel/falseLabel/lookupTargetEntity are still ignored
    # here — Dataverse doesn't support changing any of those post-creation.
    isNew: bool
    # The existing column's real logical name — required when isNew is
    # False (how update_attribute knows which column to patch). Not the
    # same as fieldRecordId, which is the catalog row's id.
    logicalName: str | None = None


class UpdateTemplatePayload(BaseModel):
    configurationId: str  # simform_reportconfiguration GUID — the SPA sets
    # simform_displaystatus to Modifying on this record itself before firing
    # this call unawaited; add_fields_to_table reverts it back to Active
    # (success or failure) once done, using this id.
    tableLogicalName: str  # the table this template's real Dataverse table
    # already resolves to — from simform_templatelogicalname.
    solutionUniqueName: str | None = None
    fields: list[UpdateFieldPayload] = []


@app.post("/api/update-template")
def update_template(payload: UpdateTemplatePayload):
    """Adds new fields to (and patches basic properties of existing fields
    on) a template whose real Dataverse table already exists. Like
    /api/create-template, the SPA no longer awaits this — it navigates away
    right after setting the configuration's status to Modifying, and this
    endpoint's job (via add_fields_to_table's finally block) is to put that
    status back to Active when the real work finishes, whatever the
    outcome."""
    try:
        body = payload.model_dump(exclude_none=True)
        return add_fields_to_table(body)
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err
    except ValueError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err


@app.get("/api/tables")
def get_tables():
    try:
        return {"tables": list_tables()}
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err


@app.get("/api/tables/{logical_name}/fields")
def get_table_fields(logical_name: str):
    try:
        return {"fields": list_table_fields(logical_name.strip().lower())}
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err


@app.get("/api/tables/{logical_name}/fields/{field_logical_name}/detail")
def get_field_detail(logical_name: str, field_logical_name: str):
    """One field's full type-specific config (Min/Max, Precision, Options
    with real Values, …) — a separate, richer endpoint from the lightweight
    /fields list above (which stays untouched, used by the unrelated Lookup
    target-entity field picker). Used when loading an existing template for
    editing, so the Edit view can show a field's current settings instead of
    guessing — see loadTemplateForEdit in the SPA."""
    try:
        cfg = load_config()
        if not cfg:
            raise DataverseError("Backend is missing Dataverse environment variables.")
        token = get_token(cfg)
        return get_attribute_detail(
            cfg,
            token,
            logical_name.strip().lower(),
            field_logical_name.strip().lower(),
        )
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err


class EnsureWebApiAccessPayload(BaseModel):
    entityLogicalName: str
    entityDisplayName: str
    fields: list[str] = []
    # The CRUD/association flags ensure_webapi_access has always accepted but
    # this endpoint previously couldn't express — it hardcoded its defaults,
    # so the one thing it could never do was grant Append/AppendTo. That gap
    # is exactly what left simform_materialloaddetails with AppendTo=False and
    # produced "You don't have permission to associate or disassociate table
    # X to Y" at submit time: writing a lookup at row-creation is an
    # association, gated on Append (the row being created) AND AppendTo (the
    # row being pointed at), separately from plain Create.
    #
    # Flags are MERGED with whatever the permission already has (see
    # ensure_webapi_access) — passing false never revokes anything, so this
    # stays safe to re-run.
    create: bool = True
    read: bool = True
    write: bool = True
    append: bool = False
    appendto: bool = False
    # Binds the "Authenticated Users" role alongside Administrators (always
    # bound). Needed for anything a signed-in portal user touches directly —
    # this site has no dedicated Service Provider role.
    bindAuthenticatedUsers: bool = False


@app.post("/api/ensure-webapi-access")
def post_ensure_webapi_access(payload: EnsureWebApiAccessPayload):
    """Manual/one-off version of what create-template now does automatically
    for every table it creates — used to bootstrap access for a table that
    already existed before this endpoint did (e.g. the catalog tables
    themselves), to re-run after hand-adding a column outside this app, or to
    grant Append/AppendTo on a pair of tables joined by a Lookup that was
    created directly against the Metadata API rather than through
    /api/create-template (which wires those flags up itself)."""
    try:
        cfg = load_config()
        if not cfg:
            raise DataverseError("Backend is missing Dataverse environment variables.")
        token = get_token(cfg)
        extra_roles = (
            [_lookup_webrole_id(cfg, token, "Authenticated Users")]
            if payload.bindAuthenticatedUsers
            else []
        )
        ensure_webapi_access(
            cfg,
            token,
            payload.entityLogicalName.strip().lower(),
            payload.entityDisplayName,
            payload.fields,
            create=payload.create,
            read=payload.read,
            write=payload.write,
            append=payload.append,
            appendto=payload.appendto,
            extra_webrole_ids=extra_roles,
        )
        return {"success": True}
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err


class EnsureGridTableLookupPayload(BaseModel):
    gridTableLogicalName: str
    templateTableLogicalName: str


@app.post("/api/ensure-grid-table-lookup")
def post_ensure_grid_table_lookup(payload: EnsureGridTableLookupPayload):
    """Called once per (grid table, template table) pair the moment an
    admin picks a real grid table for a Grid Table field and saves —
    provisions the real Lookup column that field's rows bind through at
    submit time (see grid-submissions.ts's writeGridRows), plus the
    Append/AppendTo Web API access both sides need for that bind to
    actually work. Idempotent — safe to call again on every subsequent save
    of the same field."""
    try:
        cfg = load_config()
        if not cfg:
            raise DataverseError("Backend is missing Dataverse environment variables.")
        token = get_token(cfg)
        lookup_schema_name = ensure_grid_table_lookup(
            cfg,
            token,
            payload.gridTableLogicalName.strip().lower(),
            payload.templateTableLogicalName.strip().lower(),
        )
        return {"lookupSchemaName": lookup_schema_name}
    except DataverseError as err:
        raise HTTPException(status_code=400, detail={"error": str(err)}) from err

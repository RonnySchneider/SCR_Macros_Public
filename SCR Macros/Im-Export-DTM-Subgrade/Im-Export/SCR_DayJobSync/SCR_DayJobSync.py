#   GNU GPLv3
#   <this is an add-on Script/Macro for the geospatial software "Trimble Business Center" aka TBC>
#   <you'll need at least the "Survey Advanced" licence of TBC in order to run this script>
#	<see the ToolTip section below for a brief explanation what the script does>
#	<see the Help-Files for more details>
#   Copyright (C) 2026 Ronny Schneider
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <https://www.gnu.org/licenses/>

from System.Collections.Generic import List, IEnumerable # import here, otherwise there is a weird issue with Count and Add for lists
import os
exec(open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())


class TrimbleConnectClient(object):
    """Thin wrapper around Trimble.Vce.Services.Construction.TrimbleConnect.TrimbleConnectService,
    reusing TBC's own already-authenticated Trimble ID session - no separate OAuth flow needed. GetFileList(parent, progressBar) is the one
    primitive the service exposes for browsing: called with None it returns the region list; called with
    a region it returns the projects in that region; called with a project or folder it returns that
    folder's immediate children (one level at a time - there is no single call that returns a whole tree,
    so deep folders must be browsed into rather than eagerly flattened)."""

    def __init__(self):
        self._svc = None

    def _get_service(self):
        if self._svc is None:
            clr.AddReference("Trimble.Vce.Services.Construction")
            from Trimble.Vce.Services.TrimbleConnect import TrimbleConnectService
            svc = TrimbleConnectService()
            svc.Connect(svc.Settings, None)
            self._svc = svc
        return self._svc

    def get_regions(self):
        return list(self._get_service().GetFileList(None, None) or [])

    def get_children(self, parent):
        return list(self._get_service().GetFileList(parent, None) or [])

    def save_file_remotely(self, localFile, parentFolderOrProject):
        # ProgressBar.TBC_ProgressBar lives in the same Trimble.Vce.Interfaces.ProgressBar namespace as
        # the IProgressBarControl this method expects, so it's very likely the intended control to pass -
        # fall back to no progress bar if it turns out not to satisfy the interface
        try:
            self._get_service().SaveFileRemotely(localFile, parentFolderOrProject, ProgressBar.TBC_ProgressBar)
        except Exception:
            self._get_service().SaveFileRemotely(localFile, parentFolderOrProject, None)

    def download_file(self, remoteFile, localFilePath, progressBar=None):
        # remoteFile is the IFileInformation node exactly as returned by GetFileList - no id lookup
        # needed. The ID-based DownloadFile/DownloadFileAsync(path, fileId, token) API looked like the
        # obvious choice but needs a RemoteFileId that browsed files don't carry (that interface appears
        # to be for TBC's internal "referenced file" tracking, not general Trimble Connect browsing) -
        # it silently no-oped (no exception, no file) and even deadlocked the UI thread once, presumably
        # while stuck retrying with a bad/empty id. BeginDownloadFile/EndDownloadFile is the plain
        # IAsyncResult (APM) pattern that takes the file node directly, and EndDownloadFile blocks on a
        # wait handle rather than an awaited continuation, so it's safe to call synchronously.
        svc = self._get_service()
        asyncResult = svc.BeginDownloadFile(localFilePath, remoteFile, None, None, progressBar)
        return svc.EndDownloadFile(asyncResult)

    def create_folder(self, folderName, parentFolder):
        # returns True if the folder was created OR already existed - CreateFolder is its own
        # existence check, so callers don't need to look before creating
        return self._get_service().CreateFolder(folderName, parentFolder)

    def remove_file(self, fileName, parentFolder):
        return self._get_service().RemoveFile(fileName, parentFolder, None)

    def remove_folder(self, folderName, parentFolder):
        return self._get_service().RemoveFolder(folderName, parentFolder, None)

    def get_or_create_child_folder(self, parentFolder, folderName):
        # there is no API that hands back the folder node CreateFolder just created/confirmed (it only
        # returns a bool), so after ensuring it exists we re-list parentFolder's children to find it
        self.create_folder(folderName, parentFolder)
        for child in self.get_children(parentFolder):
            if child.IsFolder and str(child.FileName) == folderName:
                return child
        raise Exception("Trimble Connect folder '" + folderName + "' could not be created or found.")

    def get_access_token(self):
        # TrimbleConnectService.AuthenticationInfo is documented only as "authentication information...
        # usually acquired by signing in", with no public type/shape - but Connect() takes a
        # Trimble.Vce.Services.Auth.TrimbleIdAuthInfo, which does publicly expose an AccessToken (a TID
        # bearer JWT), so AuthenticationInfo is almost certainly that same object. Returns None (rather
        # than raising) on any mismatch, so callers can treat "no token" as "Web API not available here".
        try:
            authInfo = self._get_service().AuthenticationInfo
            return authInfo.AccessToken if authInfo is not None else None
        except Exception:
            return None


def trimble_connect_region_base_url(remoteItem):
    # confirmed via live diagnostic: a browsed item's RootOriginBaseURL is a region-specific, protocol-
    # relative host (e.g. "//app32.connect.trimble.com" for this account's region) - NOT the generic
    # "app.connect.trimble.com" gateway the .NET SDK docs quote as the default production endpoint, so
    # calls need to go to the item's own shard rather than the generic host
    rootOrigin = getattr(remoteItem, "RootOriginBaseURL", None)
    if not rootOrigin:
        return None
    return "https:" + rootOrigin + "/tc/api/2.0"


class TrimbleConnectWebApi(object):
    """Thin wrapper around the real Trimble Connect REST API, used only for the one operation
    TrimbleConnectService has no primitive for: actually moving a file or folder to a different parent
    (PATCH .../{id} with a "parentId" body, confirmed against Trimble's public tcps/2.0 API spec and, for
    the id shape, against a live item dump - RemoteFileId matches that item's own "ID" field exactly).
    Everything else (browsing, download, upload, create/remove) goes through TrimbleConnectClient/
    TrimbleConnectService instead, since that's the one already proven to work with TBC's own signed-in
    session."""

    def __init__(self, accessToken, baseUrl):
        self.accessToken = accessToken
        self.baseUrl = baseUrl

    def _patch(self, path, bodyDict):
        request = HttpWebRequest.Create(self.baseUrl + path)
        request.Method = "PATCH"
        request.ContentType = "application/json"
        request.Accept = "application/json"
        request.Headers.Add("Authorization", "Bearer " + self.accessToken)

        payload = Encoding.UTF8.GetBytes(json.dumps(bodyDict))
        request.ContentLength = payload.Length
        stream = request.GetRequestStream()
        stream.Write(payload, 0, payload.Length)
        stream.Close()

        try:
            request.GetResponse().Close()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("Trimble Connect Web API PATCH " + path + " failed: " + str(ex.Message) + " " + detail)

    def move_file(self, fileId, newParentId):
        self._patch("/files/" + str(fileId), {"parentId": str(newParentId)})

    def move_folder(self, folderId, newParentId):
        self._patch("/folders/" + str(folderId), {"parentId": str(newParentId)})

    def _get(self, path):
        request = HttpWebRequest.Create(self.baseUrl + path)
        request.Method = "GET"
        request.Accept = "application/json"
        request.Headers.Add("Authorization", "Bearer " + self.accessToken)
        try:
            response = request.GetResponse()
            with StreamReader(response.GetResponseStream(), Encoding.UTF8) as sr:
                return sr.ReadToEnd()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("Trimble Connect Web API GET " + path + " failed: " + str(ex.Message) + " " + detail)

    def get_regions_raw(self):
        # ConnectedWorkspace.dll's own ConnectProjectInfoService matches regions by ".Location" (per its
        # decompiled RegionInfoResponse model: {"tc-api": ..., "location": ...}) called against the MASTER
        # region endpoint per Trimble's own docs, not a per-project one - hitting the real endpoint
        # directly sidesteps guessing the region code from TBC's own file-tree display labels entirely
        return self._get("/regions")

    def get_trn_region_for_origin(self, originHost):
        # Trimble's own Jobs Service Swagger docs give the authoritative example
        # trn:connect:projects:northAmerica:3AWPbPuxFaE - "northAmerica" is camelCase, matching this
        # /regions endpoint's "location" field exactly ("northAmerica", "europe", "australia", ...), not
        # TBC's own display text ("North America"/"Australia") and not the short region/trnRegion codes
        # ("na"/"us", "ap-au") tried earlier - matching by origin host avoids hardcoding this per region
        regions = json.loads(self.get_regions_raw())
        normalizedTarget = originHost.strip("/").lower()
        for r in regions:
            normalizedOrigin = str(r.get("origin", "")).strip("/").lower()
            if normalizedOrigin == normalizedTarget:
                return r.get("location")
        raise Exception("No matching region found in /regions for origin: " + originHost)



_SURVEY_JOBS_API_REGION_PREFIXES = {
    # region "location" string (from the /regions endpoint, e.g. via get_trn_region_for_origin) -> the
    # regional host prefix used by the Maps API that backs Connect's Field Data Map Viewer. Only "australia"
    # is confirmed so far (from a live DevTools capture of that viewer app hitting
    # au.api.maps.trimblegeospatial.com) - other regions are left unmapped rather than guessed, so
    # SurveyJobsApi raises instead of silently hitting a made-up host.
    "australia": "au",
}


class SurveyJobsApi(object):
    """Talks to the Trimble Geospatial Maps API (v2) that backs Connect's Field Data Map Viewer - the one
    place a survey job's status can be changed. Discovered via a live DevTools capture of that viewer app
    (viewer.maps.trimblegeospatial.com), not from any public Trimble docs - EXPERIMENTAL. Two things are
    unconfirmed and could make every call here fail: whether TBC's own Trimble Connect access token
    (the only token source available to a macro) actually carries the "MapsAPICloud" scope this API's
    Swagger implies it needs, and the region->host-prefix mapping above beyond "australia"."""

    STATUSES = ("NEW", "IN_PROGRESS", "RESOLVED", "CLOSED")  # RESOLVED = "field work complete" in the UI

    def __init__(self, accessToken, regionLocation):
        prefix = _SURVEY_JOBS_API_REGION_PREFIXES.get(regionLocation)
        if not prefix:
            raise Exception("No known Maps API host prefix for region '" + str(regionLocation) + "'")
        self.accessToken = accessToken
        self.baseUrl = "https://" + prefix + ".api.maps.trimblegeospatial.com"

    def _job_url(self, projectId, jobId):
        # the captured request's job id was a TRN (trn:jobs:job:<uuid>) - accept a bare id too and wrap it,
        # in case ConnectedWorkspaceClient's JobId/Id field turns out not to already be one
        jobTrn = str(jobId) if str(jobId).startswith("trn:") else "trn:jobs:job:" + str(jobId)
        return self.baseUrl + "/projects/" + str(projectId) + "/surveyjobs/" + jobTrn

    def _get_etag(self, projectId, jobId):
        # PATCH requires the job's current ETag as If-Match (optimistic concurrency) - only obtainable via
        # a GET first, there's no way to set a job's status blind
        request = HttpWebRequest.Create(self._job_url(projectId, jobId))
        request.Method = "GET"
        request.Accept = "application/json"
        request.Headers.Add("Authorization", "Bearer " + self.accessToken)
        try:
            response = request.GetResponse()
            try:
                return response.Headers["ETag"]
            finally:
                response.Close()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("GET survey job failed: " + str(ex.Message) + " " + detail)

    def set_status(self, projectId, jobId, status):
        if status not in self.STATUSES:
            raise Exception("Unknown survey job status: " + str(status))

        etag = self._get_etag(projectId, jobId)
        if not etag:
            raise Exception("Survey job has no ETag to send as If-Match")

        request = HttpWebRequest.Create(self._job_url(projectId, jobId))
        request.Method = "PATCH"
        request.ContentType = "application/json"
        request.Accept = "application/json"
        request.Headers.Add("Authorization", "Bearer " + self.accessToken)
        request.Headers.Add("If-Match", etag)
        request.Headers.Add("X-Field-Data-Viewer-Workspace-Type", "Survey")

        payload = Encoding.UTF8.GetBytes(json.dumps({"status": status}))
        request.ContentLength = payload.Length
        stream = request.GetRequestStream()
        stream.Write(payload, 0, payload.Length)
        stream.Close()

        try:
            request.GetResponse().Close()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("PATCH survey job status failed: " + str(ex.Message) + " " + detail)


def _dump_exception(netEx, message):
    # Exception.ToString() is .NET's own full recursive dump (every nested "---> " exception plus stack
    # traces) - written to a file since it can be long, with just the top message surfaced in the UI
    dumpPath = r"C:\temp\SCR_DayJobSync_fielddata_error.txt"
    try:
        folder = os.path.dirname(dumpPath)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with open(dumpPath, "w") as f:
            f.write(netEx.ToString())
    except Exception:
        pass
    return Exception(message + " - full dump written to " + dumpPath)


def _download_url_to_file(url, destPath):
    # Field Data attachment DownloadUrl values are pre-signed blob-storage links (confirmed by the
    # attachment listing already working without needing our own Bearer token attached here) - a plain
    # anonymous GET is enough, no Authorization header needed or wanted
    WebClient().DownloadFile(url, destPath)


def _sync_field_data_job_with_connect(token, projectTrn, jobId):
    # EXPERIMENTAL - the jobs-service Swagger docs describe this action as "Request Jobs to be synced
    # with their external source (eg Trimble Connect)". If that pushes the job's files into the project's
    # normal Connect folder structure, the whole attachment-download path above becomes unnecessary - just
    # call this and then browse Connect normally. Untested against a live call yet, so failures here are
    # swallowed by the caller rather than treated as a sync failure.
    # NOT percent-encoded, deliberately - matches the decompiled DataCollectionWorkspaceClient's own proven
    # working request-building, which concatenates its TRN into the URL raw (literal colons, unescaped)
    url = ("https://cloud.api.trimble.com/geospatial/jobs-service/1.0/projects/"
           + projectTrn + "/jobs/_sync/" + str(jobId))
    request = HttpWebRequest.Create(url)
    request.Method = "POST"
    request.Accept = "application/json"
    request.ContentLength = 0
    request.Headers.Add("Trimble-FieldSystems-ClientName", "SCR_DayJobSync")
    request.Headers.Add("Authorization", "Bearer " + token)
    try:
        response = request.GetResponse()
        with StreamReader(response.GetResponseStream(), Encoding.UTF8) as sr:
            return sr.ReadToEnd()
    except WebException as ex:
        detail = ""
        if ex.Response is not None:
            with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                detail = sr.ReadToEnd()
        raise Exception(str(ex.Message) + " " + detail)


def _await_task(task, timeoutMs=20000):
    # IronPython has no async/await - .NET Task objects are waited on synchronously instead. Task.Wait()
    # does NOT just return False on failure the way its bool return type suggests - it returns False only
    # on a genuine timeout, but THROWS the AggregateException directly the moment the task faults. The
    # previous version checked task.IsFaulted AFTER this call, which was dead code: the exception always
    # escaped one line earlier, uncaught, before that check could ever run.
    try:
        completedInTime = task.Wait(timeoutMs)
    except Exception as ex:
        # a caught .NET exception can come through IronPython as a wrapper whose real CLR exception
        # lives in .clsException rather than being directly accessible via .Message/.ToString() (same
        # gotcha SCR_CloudRevise's _describe_dotnet_exception works around) - which is exactly why the
        # dump came back as just the bare AggregateException phrase with no further detail
        netEx = getattr(ex, "clsException", ex)
        raise _dump_exception(netEx, netEx.Message)

    if not completedInTime:
        raise Exception("Timed out waiting for a response")
    if task.IsFaulted:
        netEx = task.Exception
        raise _dump_exception(netEx, netEx.Message) if netEx is not None else Exception("Task faulted with no exception info")

    # task.Result via IronPython's normal dynamic attribute lookup throws "object has no attribute
    # 'Result'" on this particular closed generic (Task<ConnectedWorkspaceResponse<Exception,
    # ProjectJobPageDataResponse>>) - its generic argument comes from an assembly loaded via
    # Assembly.LoadFrom rather than clr.AddReference, which IronPython's dynamic binder apparently
    # can't resolve members against even though the property genuinely exists. Plain .NET reflection
    # bypasses that binder entirely and works regardless.
    resultProperty = task.GetType().GetProperty("Result")
    return resultProperty.GetValue(task, None)


def _get_clr_property(obj, name, default=None):
    # same dynamic-binder gap as task.Result above - plain getattr() on these Contracts DTOs (e.g.
    # ConnectedWorkspaceResponse<Exception, ProjectJobPageDataResponse>, ProjectJobPageDataResponse)
    # silently returns the fallback default instead of raising AttributeError, since IronPython can't
    # resolve the property dynamically but getattr's 3-arg form swallows that as if it were absent.
    # Reflection sidesteps the binder and returns the real value.
    if obj is None:
        return default
    propInfo = obj.GetType().GetProperty(name)
    if propInfo is None:
        return default
    return propInfo.GetValue(obj, None)


def _describe_failed_response(response):
    # ConnectedWorkspaceResponse<Exception, ProjectJobPageDataResponse> - IsSuccess is False here, so
    # dump every public property reflection finds (Error/Exception/Message, whatever this build actually
    # calls it) rather than guessing a single name, so the real reason (auth, bad request shape, etc.)
    # is visible instead of a bare "reported failure"
    try:
        parts = []
        for propInfo in response.GetType().GetProperties():
            try:
                value = propInfo.GetValue(response, None)
            except Exception as ex:
                value = "<error reading property: " + str(ex) + ">"
            parts.append(propInfo.Name + "=" + str(value))
        return "; ".join(parts)
    except Exception as ex:
        return "(could not introspect response: " + str(ex) + ")"


def _find_field_value_by_type_name(obj, typeFullName):
    # walks an object's declared instance fields (public and private) looking for the one whose static
    # field type matches typeFullName - used to pull TBC's already-authenticated IAuthenticator out of
    # TrimbleConnectService without depending on its exact (compiler-generated, single-letter) field name
    bf = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public
    for f in obj.GetType().GetFields(bf):
        if f.FieldType.FullName == typeFullName:
            return f.GetValue(obj)
    return None


class ConnectedWorkspaceClient(object):
    """Access to Trimble Connect's "Field Data" section via Trimble.Vce.ConnectedWorkspace.dll - the
    actual client library TBC itself ships and dependency-injects for this exact feature (confirmed: it's
    listed under DependencyInjectorAssemblies in TrimbleBusinessCenter.appsettings*.json, and its model
    classes - ProjectJobStatus New/InProgress/FieldWorkComplete/Done, ProjectJobAttachmentRecord - match
    Connect's Field Data UI exactly). A macro can't reach into TBC's own DI container, so the handful of
    things DataCollectionWorkspaceClient's constructor needs are built by hand instead:
      - the real API base URL - not embedded as a literal anywhere in the DLL, but present as
        "DataCollectionWorkspaceUrl" in the same appsettings JSON file TBC itself loads at startup
      - TBC's own already-authenticated IAuthenticator, pulled off the live TrimbleConnectService via
        reflection (see _find_field_value_by_type_name)
      - a plain HttpClient, and a minimal IConnectedWorkspaceConfiguration implemented here directly,
        since TBC's own concrete implementation only has a private setter (DI-only, no way to construct
        it standalone with a chosen URL)
    Confirmed working against a live project (jobs list + attachment download URLs)."""

    def __init__(self, tcServiceSource):
        self._tcServiceSource = tcServiceSource  # callable returning a connected TrimbleConnectService
        self._lastClient = None
        self._lastInterfaceType = None

    def _get_install_dir(self):
        import System.Diagnostics
        exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName
        return os.path.dirname(exePath)

    def _find_appsettings_url(self):
        installDir = self._get_install_dir()
        for fileName in ("TrimbleBusinessCenter.appsettings.Release.json", "TrimbleBusinessCenter.appsettings.json"):
            path = os.path.join(installDir, fileName)
            if not os.path.isfile(path):
                continue
            with open(path, "r") as f:
                text = f.read()
            # this settings file has "//" line comments, so it isn't strict JSON - a direct regex pull
            # of the one key we need sidesteps writing/finding a JSON-with-comments parser for it
            match = re.search(r'"DataCollectionWorkspaceUrl"\s*:\s*"([^"]+)"', text)
            if match:
                return match.group(1)
        raise Exception("Could not find DataCollectionWorkspaceUrl in TBC's own appsettings file")

    def _reflect_type(self, assembly, simpleName):
        # a full-namespace GetType(name) lookup ("Trimble.Vce.ConnectedWorkspace.IConnectedWorkspaceConfiguration")
        # came back null despite that exact path appearing in a decompiler's namespace tree - most likely
        # it's a nested type (reflection separates those with "+", not ".") or something else about its
        # real full name doesn't match what a namespace tree view implies. Searching all of the assembly's
        # types by simple Name sidesteps needing the exact full name at all.
        matches = [t for t in assembly.GetTypes() if t.Name == simpleName]
        if not matches:
            raise Exception("Type not found in Trimble.Vce.ConnectedWorkspace.dll: " + simpleName)
        if len(matches) > 1:
            raise Exception("Multiple types named '" + simpleName + "' found: " + ", ".join(t.FullName for t in matches))
        return clr.GetPythonType(matches[0])

    def _create_instance(self, pythonType, *args):
        # calling a type obtained via clr.GetPythonType(reflectedType) directly as a constructor doesn't
        # carry proper multi-overload binding the way a statically-imported class does, and even
        # Activator.CreateInstance's default binder failed to match a same-arg-count constructor on this
        # build's actual (apparently refactored/renamed-namespace) type - so constructors are matched by
        # hand here, by parameter count, and invoked directly via ConstructorInfo.Invoke. If nothing
        # matches, every available constructor's real parameter types are reported, since this build's
        # actual signatures have already proven to differ from what the SDK's XML docs say.
        clrType = clr.GetClrType(pythonType)
        # GetConstructors() with no flags only returns PUBLIC ones - came back empty, so this "Contracts"
        # DTO's constructor (if any) is internal/non-public; searching non-public too and invoking
        # directly via ConstructorInfo.Invoke works regardless of accessibility
        bf = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance
        ctors = clrType.GetConstructors(bf)
        matching = [c for c in ctors if len(c.GetParameters()) == len(args)]
        if not matching:
            signatures = "; ".join("(" + ", ".join(p.ParameterType.Name for p in c.GetParameters()) + ")" for c in ctors)
            kind = "interface" if clrType.IsInterface else "static class" if (clrType.IsAbstract and clrType.IsSealed) else "abstract class" if clrType.IsAbstract else "class"
            allBf = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static
            memberNames = sorted(set(m.Name for m in clrType.GetMembers(allBf)))
            raise Exception(
                clrType.FullName + " (" + kind + ", generic=" + str(clrType.IsGenericType)
                + ", genericDef=" + str(clrType.IsGenericTypeDefinition) + ", assembly=" + clrType.Assembly.GetName().Name
                + ") has no constructor taking " + str(len(args)) + " argument(s). Ctors available: " + (signatures or "none")
                + ". All members (" + str(len(memberNames)) + "): " + ", ".join(memberNames))
        import System.Reflection
        try:
            return matching[0].Invoke(Array[object](list(args)))
        except System.Reflection.TargetInvocationException as ex:
            inner = ex.InnerException
            raise Exception((inner.GetType().FullName + ": " + inner.Message) if inner is not None else str(ex))

    def _construct_via_properties(self, pythonType, **propertyValues):
        # a live member dump revealed ConnectProjectInfo has a plain 0-arg constructor plus full
        # get/set properties (Id, OriginUrl, Region) with private backing fields - this build's
        # "Contracts" DTOs are built via property assignment after construction, not via a
        # positional constructor the way the (now outdated) SDK XML docs described. OriginUrl turned
        # out to be a System.Uri, not a plain string ("expected Uri, got str") - coercing by the
        # property's real reflected type here (Uri, and enums too) avoids chasing each one individually.
        instance = self._create_instance(pythonType)
        clrType = clr.GetClrType(pythonType)
        for name, value in propertyValues.items():
            propInfo = clrType.GetProperty(name)
            if propInfo is not None and isinstance(value, str):
                propType = propInfo.PropertyType
                if propType.FullName == "System.Uri":
                    value = Uri(value) if value else None
                elif propType.IsEnum and value:
                    value = System.Enum.Parse(propType, value, True)
            setattr(instance, name, value)
        return instance

    def _load_types(self):
        # "from Trimble.Vce.ConnectedWorkspace import X" reports a generic "No module named
        # ConnectedWorkspace" no matter what (tried bare clr.AddReference, AddReferenceToFileAndPath, and
        # a pre-load via Assembly.LoadFrom - identical message every time, even across a full TBC
        # restart), so IronPython's import-by-namespace resolution for this specific assembly can't be
        # trusted from here for reasons that aren't pinned down. Side-stepping that entirely: load the
        # assembly and resolve every type via plain .NET reflection instead, then convert each to a
        # usable Python type with clr.GetPythonType() - this only depends on Assembly.LoadFrom/GetType,
        # never on the "from X import Y" statement.
        installDir = self._get_install_dir()
        dllPath = os.path.join(installDir, "Trimble.Vce.ConnectedWorkspace.dll")
        if not os.path.isfile(dllPath):
            raise Exception("Trimble.Vce.ConnectedWorkspace.dll not found next to TBC's exe: " + dllPath)

        import System.Reflection
        assembly = System.Reflection.Assembly.LoadFrom(dllPath)
        try:
            assembly.GetTypes()
        except System.Reflection.ReflectionTypeLoadException as ex:
            details = "; ".join(str(le) for le in ex.LoaderExceptions if le is not None)
            raise Exception("Trimble.Vce.ConnectedWorkspace.dll loaded, but couldn't enumerate its types (likely a missing dependency): " + details)

        httpFactoryDllPath = os.path.join(installDir, "Microsoft.Extensions.Http.dll")
        if not os.path.isfile(httpFactoryDllPath):
            raise Exception("Microsoft.Extensions.Http.dll not found next to TBC's exe: " + httpFactoryDllPath)
        httpFactoryAssembly = System.Reflection.Assembly.LoadFrom(httpFactoryDllPath)

        sdkInterfacesDllPath = os.path.join(installDir, "Trimble.Sdk.Interfaces.dll")
        if not os.path.isfile(sdkInterfacesDllPath):
            raise Exception("Trimble.Sdk.Interfaces.dll not found next to TBC's exe: " + sdkInterfacesDllPath)
        sdkInterfacesAssembly = System.Reflection.Assembly.LoadFrom(sdkInterfacesDllPath)

        return {
            "IConnectedWorkspaceConfiguration": self._reflect_type(assembly, "IConnectedWorkspaceConfiguration"),
            "DataCollectionWorkspaceClient": self._reflect_type(assembly, "DataCollectionWorkspaceClient"),
            "ConnectProjectInfo": self._reflect_type(assembly, "ConnectProjectInfo"),
            "ProjectJobPaginationRequest": self._reflect_type(assembly, "ProjectJobPaginationRequest"),
            "IHttpClientFactory": self._reflect_type(httpFactoryAssembly, "IHttpClientFactory"),
            "ConnectProjectTransactionPrefixServiceDelegate": self._reflect_type(assembly, "ConnectProjectTransactionPrefixServiceDelegate"),
            "ConnectProjectTransactionPrefixService": self._reflect_type(assembly, "ConnectProjectTransactionPrefixService"),
            "IDataCollectionWorkspaceClient": self._reflect_type(assembly, "IDataCollectionWorkspaceClient"),
            "IAnalyticsReporter": self._reflect_type(sdkInterfacesAssembly, "IAnalyticsReporter"),
        }

    def _invoke_interface_method(self, interfaceType, instance, methodName, *args):
        # "'DataCollectionWorkspaceClient' object has no attribute 'GetJobs'" on an object whose
        # constructor clearly succeeded means GetJobs is an EXPLICIT interface implementation - only
        # reachable through an IDataCollectionWorkspaceClient-typed reference, not the concrete class
        # (that's a compile-time/vtable distinction C# enforces; IronPython's normal dynamic attribute
        # lookup on the concrete instance can't see it). Reflection's MethodInfo.Invoke dispatches
        # correctly either way, bypassing that entirely.
        #
        # This runs on TBC's UI thread, and _await_task blocks that same thread synchronously on
        # Task.Wait(). The async method's continuation past its first "await" captures
        # SynchronizationContext.Current at that point (standard C# async behavior) and tries to post
        # itself back onto the UI thread to resume - which can never run while Task.Wait() is blocking
        # that very thread, so the call hung until _await_task's own timeout every time. Capture happens
        # the instant the async method starts running inside Invoke() below, so the context has to be
        # nulled out BEFORE that call (nulling it only around the later Wait() would be too late) - with
        # no captured context, the continuation resumes on a thread-pool thread instead, so it can
        # actually run while the UI thread waits.
        import System.Threading
        ifaceClrType = clr.GetClrType(interfaceType)
        method = ifaceClrType.GetMethod(methodName)
        if method is None:
            raise Exception("Method not found on " + ifaceClrType.FullName + ": " + methodName)
        originalContext = System.Threading.SynchronizationContext.Current
        System.Threading.SynchronizationContext.SetSynchronizationContext(None)
        try:
            return method.Invoke(instance, Array[object](list(args)))
        finally:
            System.Threading.SynchronizationContext.SetSynchronizationContext(originalContext)

    def _build_client(self, types, connectProjectInfo):
        clr.AddReference("System.Net.Http")
        from System.Net.Http import HttpClient

        authenticator = _find_field_value_by_type_name(self._tcServiceSource(), "Trimble.Vce.Services.Auth.IAuthenticator")
        if authenticator is None:
            raise Exception("Could not find TBC's IAuthenticator on the connected Trimble Connect service")

        url = self._find_appsettings_url()

        # check the interface's real property type instead of assuming string vs Uri again - OriginUrl
        # on ConnectProjectInfo turned out to be a Uri where the docs implied string, so the same
        # refactor may equally apply here
        configPropType = clr.GetClrType(types["IConnectedWorkspaceConfiguration"]).GetProperty("DataCollectionWorkspaceUrl").PropertyType
        urlValue = Uri(url) if configPropType.FullName == "System.Uri" else url

        class _FixedWorkspaceConfiguration(types["IConnectedWorkspaceConfiguration"]):
            def __init__(self, value):
                self._value = value

            @property
            def DataCollectionWorkspaceUrl(self):
                return self._value

        # same TID access token TrimbleConnectClient.get_access_token() already proved works - a real
        # DI-configured named HttpClient would normally get its Authorization header attached by a
        # message handler wired up alongside the factory, which our bare stub factory doesn't replicate
        accessToken = None
        try:
            accessToken = self._tcServiceSource().AuthenticationInfo.AccessToken
        except Exception:
            pass

        clr.AddReference("System.Net.Http")
        from System.Net.Http.Headers import AuthenticationHeaderValue

        # a real DI-configured named HttpClient would already have BaseAddress (from
        # DataCollectionWorkspaceUrl) and an Authorization header set - our bare HttpClient() had
        # neither, which is what an internal Trimble.Vce.ConnectedWorkspace.HttpClientExtensions
        # helper NullReferenceException'd on
        class _SimpleHttpClientFactory(types["IHttpClientFactory"]):
            def CreateClient(self, name):
                client = HttpClient()
                client.BaseAddress = Uri(url)
                if accessToken:
                    client.DefaultRequestHeaders.Authorization = AuthenticationHeaderValue("Bearer", accessToken)
                return client

        # decompiled HttpClientExtensions.ExecuteRequest calls analyticsReporter.Report(...)/
        # .ReportException(...) UNCONDITIONALLY, on both the "request failed" and "any exception" paths -
        # with no null-check. A bare None for this constructor argument is exactly what was NRE-ing,
        # likely while trying to log the real failure and masking it - a no-op stub lets the real
        # response/error actually come back instead.
        class _NoOpAnalyticsReporter(types["IAnalyticsReporter"]):
            def Report(self, *args):
                pass

            def ReportCommandAction(self, *args):
                pass

            def ReportException(self, *args):
                pass

        # this build's real constructor (found via live reflection, differs from the SDK's XML docs) is
        # (ConnectProjectInfo, IHttpClientFactory, IConnectedWorkspaceConfiguration, IAnalyticsReporter,
        # IAuthenticator, ConnectProjectTransactionPrefixServiceDelegate). A generic no-op stub for this
        # delegate (returning None, since its return type is an interface with no obvious safe default)
        # NRE'd the moment GetJobs/GetJobAttachmentDetails/etc. called .GetConnectProjectTransactionPrefix()
        # on that null. The real DLL class (ConnectProjectTransactionPrefixService, instantiated here via
        # reflection rather than reimplemented) builds "trn:connect:projects:" + Region + ":" + Id, matching
        # Trimble's own documented TRN example ("trn:connect:projects:northAmerica:3AWPbPuxFaE") - Region
        # has to be the /regions endpoint's camelCase "location" value (e.g. "australia"), not TBC's own
        # display text or any short region code, both of which get rejected as "not a valid format".
        def _real_transaction_prefix_delegate(connectProjectInfoArg):
            return self._create_instance(types["ConnectProjectTransactionPrefixService"], connectProjectInfoArg)

        transactionPrefixDelegate = types["ConnectProjectTransactionPrefixServiceDelegate"](_real_transaction_prefix_delegate)
        return self._create_instance(
            types["DataCollectionWorkspaceClient"],
            connectProjectInfo,
            _SimpleHttpClientFactory(),
            _FixedWorkspaceConfiguration(urlValue),
            _NoOpAnalyticsReporter(),
            authenticator,
            transactionPrefixDelegate,
        )

    def get_jobs(self, projectId, region, originUrl):
        types = self._load_types()
        connectProjectInfo = self._construct_via_properties(types["ConnectProjectInfo"], Id=projectId, Region=region, OriginUrl=originUrl)
        client = self._build_client(types, connectProjectInfo)

        try:
            paginationRequest = self._create_instance(types["ProjectJobPaginationRequest"], 0, 200)
        except Exception:
            # same "Contracts" refactor as ConnectProjectInfo likely applies here too - fall back to
            # property assignment if the 2-arg constructor this SDK's docs describe doesn't exist either
            paginationRequest = self._construct_via_properties(types["ProjectJobPaginationRequest"], PageIndex=0, PageSize=200)
        task = self._invoke_interface_method(types["IDataCollectionWorkspaceClient"], client, "GetJobs", paginationRequest, None)
        response = _await_task(task)
        if not _get_clr_property(response, "IsSuccess", True):
            raise Exception("GetJobs reported failure - " + _describe_failed_response(response))
        pageData = _get_clr_property(response, "ResponseData", response)
        items = _get_clr_property(pageData, "Items", pageData)

        self._lastClient = client  # reused for get_attachments - same client/project context
        self._lastInterfaceType = types["IDataCollectionWorkspaceClient"]
        return list(items) if items is not None else []

    def get_attachments(self, jobId):
        if self._lastClient is None:
            raise Exception("Load jobs first")
        # real signature is GetJobAttachmentDetails(string jobId, bool includeDownloadUrl, CancellationToken? c) -
        # includeDownloadUrl=True since we need a direct download URL for each attachment to sync files
        task = self._invoke_interface_method(self._lastInterfaceType, self._lastClient, "GetJobAttachmentDetails", jobId, True, None)
        response = _await_task(task)
        if not _get_clr_property(response, "IsSuccess", True):
            raise Exception("GetJobAttachmentDetails reported failure")
        items = _get_clr_property(response, "ResponseData", response)
        return list(items) if items is not None else []

    def get_job_details(self, jobId):
        # GetJobs (the paged list) doesn't populate RootDataFile - only this per-job call does (confirmed
        # live: RootDataFile.Id came back None from a job pulled off the GetJobs list). Needed to build
        # the "trn:jobs:connect-file:{region}:{fileId}" TRN the jobs-service's _sync action expects.
        if self._lastClient is None:
            raise Exception("Load jobs first")
        task = self._invoke_interface_method(self._lastInterfaceType, self._lastClient, "GetJobDetails", jobId, None)
        response = _await_task(task)
        if not _get_clr_property(response, "IsSuccess", True):
            raise Exception("GetJobDetails reported failure - " + _describe_failed_response(response))
        return _get_clr_property(response, "ResponseData", response)


class ExcelJobRegister(object):
    """Writes a value into the next empty row of one column in an existing Excel workbook, via late-bound
    COM automation (Excel.Application) - just needs Excel installed, no PIA reference or extra DLL. The
    workbook is left open and visible afterwards (not saved-and-closed) so the user can fill in the rest
    of that row by hand - reuses an already-running Excel/already-open copy of the workbook if there is
    one (the user's own, or one left open by a previous sync), rather than opening a second copy."""

    XL_UP = -4162

    def append_value(self, workbookPath, columnLetter, value, minimumRow=1):
        excelApp, weOwnTheApp = self._get_or_create_application()
        workbook = None
        try:
            workbook = self._find_open_workbook(excelApp, workbookPath)
            if workbook is None:
                workbook = excelApp.Workbooks.Open(workbookPath)

            worksheet = workbook.Worksheets[1]
            nextRow = max(self._find_next_empty_row(worksheet, columnLetter), minimumRow)
            cell = worksheet.Range(columnLetter + str(nextRow))
            cell.Value2 = value

            workbook.Save()

            excelApp.Visible = True
            workbook.Activate()
            worksheet.Activate()
            cell.Select()
            return nextRow
        finally:
            # deliberately no workbook.Close()/excelApp.Quit() - the whole point is to leave it open for
            # the user. Releasing our own COM references is still correct hygiene: Excel is a genuinely
            # separate process (out-of-process COM), so dropping our proxy to it doesn't close anything
            # the user can see - it just means this macro is no longer holding a handle to it.
            if workbook is not None:
                try:
                    Marshal.ReleaseComObject(workbook)
                except Exception:
                    pass
            try:
                Marshal.ReleaseComObject(excelApp)
            except Exception:
                pass
            GC.Collect()
            GC.WaitForPendingFinalizers()

    def _get_or_create_application(self):
        # (app, weOwnIt) - weOwnIt is False when we attached to an Excel the user already had running,
        # which matters only in spirit here (we never Quit() either way) but documents the intent
        try:
            return Marshal.GetActiveObject("Excel.Application"), False
        except Exception:
            excelType = Type.GetTypeFromProgID("Excel.Application")
            return Activator.CreateInstance(excelType), True

    def _find_open_workbook(self, excelApp, workbookPath):
        targetPath = os.path.normcase(os.path.abspath(workbookPath))
        for workbook in excelApp.Workbooks:
            try:
                if os.path.normcase(os.path.abspath(workbook.FullName)) == targetPath:
                    return workbook
            except Exception:
                continue
        return None

    def _find_next_empty_row(self, worksheet, columnLetter):
        # standard Ctrl+Up idiom: from the very bottom row, jump up to the last non-empty cell - if the
        # whole column is empty, that jump lands back on row 1, so that row-1 case needs its own check
        # (a non-empty row 1 means row 1 IS the last used row; an empty row 1 means nothing is used yet).
        # This already copes with gaps/leading blanks - it always finds the true last used row from the
        # bottom regardless of what's empty above it - so minimumRow (above) exists only as an explicit
        # floor for the edge case where row 1 itself is blank but shouldn't be written into (e.g. a
        # header whose cell is technically empty, like a merged cell).
        bottomRow = worksheet.Rows.Count
        lastUsedRow = worksheet.Range(columnLetter + str(bottomRow)).End(self.XL_UP).Row
        firstCellValue = worksheet.Range(columnLetter + "1").Value2
        if lastUsedRow == 1 and (firstCellValue is None or str(firstCellValue).strip() == ""):
            return 1
        return lastUsedRow + 1


def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_DayJobSync"
    cmdData.CommandName = "SCR_DayJobSync"
    cmdData.Caption = "_SCR_DayJobSync"
    #cmdData.UIForm = "SCR_DayJobSync"      # left disabled - this is a fully independent floating window, not a TBC-managed dialog
                                                        # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "0"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Update/Transfer"
        cmdData.ShortCaption = "DayJob Sync"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3
        cmdData.EnableNoProject       = True

        cmdData.Version = 1.033
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "DayJobSync"
        cmdData.ToolTipTextFormatted = "log in to Trimble Connect and browse its folder tree"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass

def Execute(cmd, currentProject, macroFileFolder, parameters):
    form = SCR_DayJobSyncDialog(currentProject, macroFileFolder).Show()
    return
    # .Show() - is non modal - you can interact with the drawing window
    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window


class SCR_DayJobSyncDialog(Window): # this inherits from the WPF Window control - a fully independent floating window
    def __init__(self, currentProject, macroFileFolder):

        with StreamReader(macroFileFolder + r"\SCR_DayJobSync.xaml") as s:
            wpf.LoadComponent(self, s)

        ElementHost.EnableModelessKeyboardInterop(self)

        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.trimbleConnectClient = TrimbleConnectClient()
        self.connectedWorkspaceClient = ConnectedWorkspaceClient(self.trimbleConnectClient._get_service)
        self.projectRoot = None
        self.folderStack = []
        self._fieldDataProjectTrn = None
        self._fieldDataRegion = None
        self._fieldDataProjectId = None

        self.folderStructureBox.Text = OptionsManager.GetString("SCR_DayJobSync.folderstructure", "{YYYY}/{YYMM}/{Job}")
        self.controllerDataBox.Text = OptionsManager.GetString("SCR_DayJobSync.controllerdataname", "Controller Data")
        self.tbcBox.Text = OptionsManager.GetString("SCR_DayJobSync.tbcname", "TBC")
        self.controllerDataEnabledCheckbox.IsChecked = OptionsManager.GetString("SCR_DayJobSync.controllerdataenabled", "True") != "False"
        self.tbcEnabledCheckbox.IsChecked = OptionsManager.GetString("SCR_DayJobSync.tbcenabled", "True") != "False"
        self.oldJobFolderBox.Text = OptionsManager.GetString("SCR_DayJobSync.oldjobfoldername", "Old Jobs")
        self.oldJobFolderEnabledCheckbox.IsChecked = OptionsManager.GetString("SCR_DayJobSync.oldjobfolderenabled", "True") != "False"
        self.localSyncFolder = OptionsManager.GetString("SCR_DayJobSync.localsyncfolder", "")
        self.update_local_folder_label()

        self.jobRegisterPath = OptionsManager.GetString("SCR_DayJobSync.jobregisterpath", "")
        self.jobRegisterColumnBox.Text = OptionsManager.GetString("SCR_DayJobSync.jobregistercolumn", "A")
        self.jobRegisterStartRowBox.Text = OptionsManager.GetString("SCR_DayJobSync.jobregisterstartrow", "1")
        self.update_job_register_label()

        try:
            fileListStar = float(OptionsManager.GetString("SCR_DayJobSync.fileliststar", "3"))
            fieldDataListStar = float(OptionsManager.GetString("SCR_DayJobSync.fielddataliststar", "2"))
            self.fileListColumn.Width = GridLength(fileListStar, GridUnitType.Star)
            self.fieldDataListColumn.Width = GridLength(fieldDataListStar, GridUnitType.Star)
        except Exception:
            pass  # keep the XAML-default split if the saved value is missing/unparsable

        self.folderStructureBox.TextChanged += self.folder_structure_changed
        self.controllerDataBox.TextChanged += self.controller_data_name_changed
        self.tbcBox.TextChanged += self.tbc_name_changed
        self.controllerDataEnabledCheckbox.Checked += self.controller_data_enabled_changed
        self.controllerDataEnabledCheckbox.Unchecked += self.controller_data_enabled_changed
        self.tbcEnabledCheckbox.Checked += self.tbc_enabled_changed
        self.tbcEnabledCheckbox.Unchecked += self.tbc_enabled_changed
        self.oldJobFolderBox.TextChanged += self.old_job_folder_name_changed
        self.oldJobFolderEnabledCheckbox.Checked += self.old_job_folder_enabled_changed
        self.oldJobFolderEnabledCheckbox.Unchecked += self.old_job_folder_enabled_changed
        self.browseLocalFolderBtn.Click += self.browse_local_folder_clicked
        self.browseJobRegisterBtn.Click += self.browse_job_register_clicked
        self.jobRegisterColumnBox.TextChanged += self.job_register_column_changed
        self.jobRegisterStartRowBox.TextChanged += self.job_register_start_row_changed

        self.reloadBtn.Click += self.reload_regions_clicked
        self.regionCombo.SelectionChanged += self.region_selection_changed
        self.projectCombo.SelectionChanged += self.project_selection_changed
        self.upBtn.Click += self.up_clicked
        self.fileList.MouseDoubleClick += self.file_list_double_click
        self.fileList.SelectionChanged += self.file_list_selection_changed

        self.wmReloadBtn.Click += self.wm_reload_clicked
        self.fieldDataList.SelectionChanged += self.field_data_list_selection_changed

        self.syncBtn.Click += self.sync_clicked
        self.helpBtn.Click += self.help_clicked

        self.Loaded += self.SetDefaultOptions
        self.Loaded += self.start_initial_load
        self.Closing += self.SaveOptions


    # ---------- local sync folder / subfolder-name settings ----------

    def update_local_folder_label(self):
        self.localFolderLabel.Text = self.localSyncFolder if self.localSyncFolder else "(not set)"

    def browse_local_folder_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        dlg.Description = "Choose the local sync folder"
        if self.localSyncFolder and os.path.isdir(self.localSyncFolder):
            dlg.SelectedPath = self.localSyncFolder

        if dlg.ShowDialog() == DialogResult.OK:
            self.localSyncFolder = dlg.SelectedPath
            OptionsManager.SetValue("SCR_DayJobSync.localsyncfolder", self.localSyncFolder)
            self.update_local_folder_label()

    def folder_structure_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.folderstructure", self.folderStructureBox.Text)

    def controller_data_name_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.controllerdataname", self.controllerDataBox.Text)

    def tbc_name_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.tbcname", self.tbcBox.Text)

    def controller_data_enabled_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.controllerdataenabled", str(bool(self.controllerDataEnabledCheckbox.IsChecked)))

    def tbc_enabled_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.tbcenabled", str(bool(self.tbcEnabledCheckbox.IsChecked)))

    def old_job_folder_name_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.oldjobfoldername", self.oldJobFolderBox.Text)

    def old_job_folder_enabled_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.oldjobfolderenabled", str(bool(self.oldJobFolderEnabledCheckbox.IsChecked)))


    # ---------- job register (Excel log) settings ----------

    def update_job_register_label(self):
        self.jobRegisterLabel.Text = self.jobRegisterPath if self.jobRegisterPath else "(not set)"

    def browse_job_register_clicked(self, sender, e):
        dlg = OpenFileDialog()
        dlg.Title = "Choose the job register Excel workbook"
        dlg.Filter = "Excel Workbooks (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*"
        if self.jobRegisterPath and os.path.isfile(self.jobRegisterPath):
            dlg.InitialDirectory = os.path.dirname(self.jobRegisterPath)
            dlg.FileName = os.path.basename(self.jobRegisterPath)

        if dlg.ShowDialog() == DialogResult.OK:
            self.jobRegisterPath = dlg.FileName
            OptionsManager.SetValue("SCR_DayJobSync.jobregisterpath", self.jobRegisterPath)
            self.update_job_register_label()

    def job_register_column_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.jobregistercolumn", self.jobRegisterColumnBox.Text)

    def job_register_start_row_changed(self, sender, e):
        OptionsManager.SetValue("SCR_DayJobSync.jobregisterstartrow", self.jobRegisterStartRowBox.Text)


    # ---------- window position/size ----------

    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadWindowState(self, "SCR_DayJobSync", default_width=340, default_height=460)

    def start_initial_load(self, sender, e):
        # deferred (Background priority, below rendering) so the window is already visible and painted
        # before the Trimble Connect connection attempt starts - otherwise the whole window can sit
        # blank/frozen for a while with no indication the macro has actually started
        self.Dispatcher.BeginInvoke(DispatcherPriority.Background, Action(lambda: self.reload_regions_clicked(None, None)))

    def SaveOptions(self, sender, e):
        SCROptions.SaveWindowState(self, "SCR_DayJobSync")
        # the GridSplitter between the two lists has no per-drag "changed" event to hook, so its position
        # (the two star-sized ColumnDefinitions' current ratio) is saved once here on Closing instead
        OptionsManager.SetValue("SCR_DayJobSync.fileliststar", str(self.fileListColumn.Width.Value))
        OptionsManager.SetValue("SCR_DayJobSync.fielddataliststar", str(self.fieldDataListColumn.Width.Value))


    # ---------- remembered combo selections ----------

    def get_saved(self, key):
        return OptionsManager.GetString("SCR_DayJobSync." + key, "")

    def save_selected(self, key, value):
        OptionsManager.SetValue("SCR_DayJobSync." + key, str(value) if value is not None else "")

    def select_preferred_by_content(self, combo, optionKey):
        # Trimble Connect regions/projects have no stable id we can persist (region.ID is always None),
        # so their display name is the only usable key
        preferred = self.get_saved(optionKey) or None

        target = None
        if preferred is not None:
            for item in combo.Items:
                if str(item.Content) == str(preferred):
                    target = item
                    break

        if target is not None:
            combo.SelectedItem = target
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0


    # ---------- Trimble Connect login + folder tree ----------
    # login is implicit: TrimbleConnectService.Connect() reuses TBC's own already-authenticated Trimble ID
    # session, so the first call that touches the service (get_regions, below) is effectively "logging in".
    # Browsing is one level at a time via double-click/Up, since GetFileList only returns one folder's
    # immediate children per call (no single "whole tree" call).

    def reload_regions_clicked(self, sender, e):
        self.error.Content = ""
        self.statusLabel.Text = "Connecting to Trimble Connect..."
        self.statusLabel.Foreground = SolidColorBrush(Colors.Red)
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

        self.regionCombo.Items.Clear()
        self.projectCombo.Items.Clear()
        self.fileList.Items.Clear()
        self.projectRoot = None
        self.folderStack = []

        try:
            regions = self.trimbleConnectClient.get_regions()
        except Exception as ex:
            self.statusLabel.Text = ""
            self.statusLabel.Foreground = SolidColorBrush(Colors.Gray)
            self.error.Content = str(ex)
            return

        for r in sorted(regions, key=lambda x: str(x.FileName)):
            item = ComboBoxItem()
            item.Content = str(r.FileName)
            item.Tag = r
            self.regionCombo.Items.Add(item)

        self.statusLabel.Text = "Connected to Trimble Connect."
        self.statusLabel.Foreground = SolidColorBrush(Colors.Green)
        self.select_preferred_by_content(self.regionCombo, "selectedregionname")

        # TEMPORARY - was for manually testing the jobs-service _sync endpoint in Swagger (see the parked
        # experiment in sync_field_data_job); left commented rather than deleted in case that's resumed.
        # try:
        #     token = self.trimbleConnectClient.get_access_token()
        #     if token:
        #         with open(r"C:\temp\SCR_DayJobSync_token_PRIVATE.txt", "w") as tf:
        #             tf.write("Bearer " + token)
        # except Exception:
        #     pass

    def region_selection_changed(self, sender, e):
        self.projectCombo.Items.Clear()
        self.fileList.Items.Clear()
        self.projectRoot = None
        self.folderStack = []
        self.error.Content = ""

        regionItem = self.regionCombo.SelectedItem
        if regionItem is None:
            return
        self.save_selected("selectedregionname", regionItem.Content)

        try:
            projects = self.trimbleConnectClient.get_children(regionItem.Tag)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for p in sorted(projects, key=lambda x: str(x.FileName)):
            item = ComboBoxItem()
            item.Content = str(p.FileName)
            item.Tag = p
            self.projectCombo.Items.Add(item)

        self.select_preferred_by_content(self.projectCombo, "selectedprojectname")

    def project_selection_changed(self, sender, e):
        self.fileList.Items.Clear()
        self.folderStack = []
        self.error.Content = ""

        projectItem = self.projectCombo.SelectedItem
        if projectItem is None:
            self.projectRoot = None
            self.tcPathLabel.Text = "/"
            return
        self.save_selected("selectedprojectname", projectItem.Content)

        self.projectRoot = projectItem.Tag
        self.restore_folder_stack()
        self.refresh_file_list()

        # Field Data jobs belong to the whole project, not to whatever folder is browsed into, so they
        # only need reloading when the project itself changes - the separate reload button (wmReloadBtn)
        # still works on its own for a manual refresh without re-selecting the project
        self.wm_reload_clicked(None, None)

    def current_folder(self):
        return self.folderStack[-1] if self.folderStack else self.projectRoot

    def save_folder_path(self):
        names = [str(f.FileName) for f in self.folderStack]
        OptionsManager.SetValue("SCR_DayJobSync.lastfolderpath", json.dumps(names))

    def restore_folder_stack(self):
        # walks the saved folder path (a list of folder names from the project root down to the last
        # folder the user browsed into) one level at a time via get_children - best effort, stopping at
        # whichever level no longer matches (e.g. a different project, or the folder was renamed/removed)
        raw = self.get_saved("lastfolderpath")
        if not raw:
            return
        try:
            names = json.loads(raw)
        except Exception:
            return

        current = self.projectRoot
        stack = []
        for name in names:
            try:
                children = self.trimbleConnectClient.get_children(current)
            except Exception:
                break
            match = None
            for c in children:
                if c.IsFolder and str(c.FileName) == name:
                    match = c
                    break
            if match is None:
                break
            stack.append(match)
            current = match

        self.folderStack = stack

    def refresh_file_list(self):
        self.fileList.Items.Clear()
        self.tcPathLabel.Text = "/" + "/".join(str(f.FileName) for f in self.folderStack)

        folder = self.current_folder()
        if folder is None:
            return

        try:
            children = self.trimbleConnectClient.get_children(folder)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for child in sorted(children, key=lambda c: (not c.IsFolder, str(c.FileName))):
            item = ListBoxItem()
            item.Content = ("[folder] " if child.IsFolder else "") + str(child.FileName)
            item.Tag = child
            self.fileList.Items.Add(item)

    def file_list_double_click(self, sender, e):
        item = self.fileList.SelectedItem
        if item is None or not item.Tag.IsFolder:
            return
        self.folderStack.append(item.Tag)
        self.save_folder_path()
        self.refresh_file_list()

    def file_list_selection_changed(self, sender, e):
        # the Sync button acts on whichever side has a selection - keeping the two mutually exclusive
        # (rather than erroring after the fact when Sync is clicked) means there's never an ambiguous
        # "both sides selected" state to resolve
        if self.fileList.SelectedItem is not None:
            self.fieldDataList.SelectedItem = None

    def up_clicked(self, sender, e):
        if self.folderStack:
            self.folderStack.pop()
            self.save_folder_path()
            self.refresh_file_list()


    # ---------- Field Data (beta) - separate ConnectedWorkspace-backed panel next to the Connect tree ----------
    # unrelated to TrimbleConnectClient above: Field Data (Trimble Access field-to-office sync jobs) lives
    # in a completely different backend that GetFileList never surfaces, reusing the SAME Connect project
    # already selected on the left (Field Data belongs to a project, not a separate account/project space
    # like the earlier WorksManager dead end was). See ConnectedWorkspaceClient's docstring for details.

    def wm_reload_clicked(self, sender, e):
        self.wmStatus.Text = "Loading Field Data jobs..."
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

        self.fieldDataList.Items.Clear()

        projectItem = self.projectCombo.SelectedItem
        if projectItem is None:
            self.wmStatus.Text = "Select a Trimble Connect project on the left first"
            return

        projectId = getattr(projectItem.Tag, "ID", None) or getattr(projectItem.Tag, "RemoteFileId", None)
        if not projectId:
            self.wmStatus.Text = "Selected project has no usable id - unexpected response shape"
            return

        rootOrigin = getattr(projectItem.Tag, "RootOriginBaseURL", None)
        if not rootOrigin:
            self.wmStatus.Text = "Selected project has no origin URL - unexpected response shape"
            return
        originUrl = "https:" + rootOrigin

        # Trimble's own Jobs Service docs give the authoritative TRN example
        # "trn:connect:projects:northAmerica:3AWPbPuxFaE" - "northAmerica" is the /regions endpoint's
        # camelCase "location" field, not TBC's own display text ("Australia") or any short region code
        try:
            token = self.trimbleConnectClient.get_access_token()
            api = TrimbleConnectWebApi(token, "https://app.connect.trimble.com/tc/api/2.0")
            region = api.get_trn_region_for_origin(rootOrigin)
        except Exception as ex:
            self.wmStatus.Text = "Could not resolve TRN region - " + str(ex)
            return

        # stashed for sync_field_data_job, which calls the jobs-service's own _sync action on this same
        # project TRN once a job's files are downloaded - see _sync_field_data_job_with_connect
        self._fieldDataProjectTrn = "trn:connect:projects:" + region + ":" + projectId
        self._fieldDataRegion = region
        self._fieldDataProjectId = projectId

        try:
            jobs = self.connectedWorkspaceClient.get_jobs(projectId, region, originUrl)
        except Exception as ex:
            self.wmStatus.Text = "Error: " + str(ex)
            return

        for job in jobs:
            item = ListBoxItem()
            item.Content = _get_clr_property(job, "Name", None) or str(job)
            item.Tag = job
            self.fieldDataList.Items.Add(item)

        self.wmStatus.Text = (str(len(jobs)) + " job(s) found") if jobs else "No Field Data jobs found for this project"

    def field_data_list_selection_changed(self, sender, e):
        if self.fieldDataList.SelectedItem is not None:
            self.fileList.SelectedItem = None


    # ---------- sync ----------

    def render_folder_structure(self, yy, mm, jobName):
        # user-defined template like "{YYYY}/{YYMM}/{Job}" - {YYYY} is assumed to belong to the 2000s
        # (20yy), fine for any realistic day-job date. Split on "/" (and "\" in case someone uses that)
        # into path segments, dropping empty ones, so os.path.join(*segments) builds the actual folders.
        # {YYMM} is replaced before the standalone {YY}/{MM} so it isn't first broken up into "{YY}{MM}"
        # by those replacements running first.
        template = self.folderStructureBox.Text.strip() or "{YYYY}/{YYMM}/{Job}"
        rendered = template.replace("{YYYY}", "20" + yy).replace("{YYMM}", yy + mm) \
                            .replace("{YY}", yy).replace("{MM}", mm).replace("{Job}", jobName)
        segments = [p for p in rendered.replace("\\", "/").split("/") if p]
        if not segments:
            raise Exception("Folder Structure resolved to nothing - check the template: " + template)
        return segments

    def sync_clicked(self, sender, e):
        self.error.Content = ""
        self.statusLabel.Text = ""
        self.statusLabel.Foreground = SolidColorBrush(Colors.Gray)

        connectItem = self.fileList.SelectedItem
        fieldDataItem = self.fieldDataList.SelectedItem
        if connectItem is not None and fieldDataItem is not None:
            self.error.Content = "A file is selected on both sides - select one or the other, not both."
            return
        if fieldDataItem is not None:
            self.sync_field_data_job(fieldDataItem.Tag)
            return
        if connectItem is None:
            self.error.Content = "Select a file in the Trimble Connect folder tree, or a Field Data job, first."
            return

        self.sync_connect_file(connectItem.Tag)

    def sync_connect_file(self, connectFile):
        if connectFile.IsFolder:
            self.error.Content = "Select a file, not a folder."
            return

        if not self.localSyncFolder or not os.path.isdir(self.localSyncFolder):
            self.error.Content = "Set a valid local sync folder first."
            return

        fileName = str(connectFile.FileName)
        match = re.match(r"^(\d{2})(\d{2})\d{2}", fileName)
        if match is None:
            self.error.Content = "Filename does not start with a 6-digit date (YYMMDD): " + fileName
            return
        yy, mm = match.group(1), match.group(2)

        controllerDataEnabled = bool(self.controllerDataEnabledCheckbox.IsChecked)
        tbcEnabled = bool(self.tbcEnabledCheckbox.IsChecked)
        controllerDataName = self.controllerDataBox.Text.strip()
        tbcName = self.tbcBox.Text.strip()
        if controllerDataEnabled and not controllerDataName:
            self.error.Content = "Set the 'Controller Data' subfolder name, or untick its checkbox."
            return
        if tbcEnabled and not tbcName:
            self.error.Content = "Set the 'TBC' subfolder name, or untick its checkbox."
            return

        folderName = os.path.splitext(fileName)[0]
        try:
            targetFolder = os.path.join(self.localSyncFolder, *self.render_folder_structure(yy, mm, folderName))
        except Exception as ex:
            self.error.Content = str(ex)
            return
        controllerDataFolder = os.path.join(targetFolder, controllerDataName) if controllerDataEnabled else targetFolder
        # if Controller Data isn't ticked, the job (and its Files subfolder) is placed directly in
        # targetFolder instead of a Controller Data subfolder - controllerDataFolder just aliases
        # targetFolder in that case, so every downstream use of it already does the right thing

        try:
            foldersToCreate = [targetFolder]
            if controllerDataEnabled:
                foldersToCreate.append(controllerDataFolder)
            if tbcEnabled:
                foldersToCreate.append(os.path.join(targetFolder, tbcName))
            for folder in foldersToCreate:
                if not os.path.isdir(folder):
                    os.makedirs(folder)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        downloadPath = os.path.join(controllerDataFolder, fileName)

        self.syncBtn.IsEnabled = False
        self.statusLabel.Text = "Downloading " + fileName + "..."
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

        try:
            self.trimbleConnectClient.download_file(connectFile, downloadPath)
        except Exception as ex:
            self.syncBtn.IsEnabled = True
            self.error.Content = "Folders created, but download failed: " + str(ex)
            return

        if not self.wait_for_file(downloadPath, expectedSize=connectFile.Size):
            self.syncBtn.IsEnabled = True
            self.error.Content = "Download reported success, but the file is not at: " + downloadPath
            return

        # a job file can have an accompanying "<name without suffix> Files" subfolder alongside it on
        # Connect (e.g. "260831SCR1.job" -> "260831SCR1 Files") holding its supporting files - if one
        # exists next to the job file, copy the whole thing into the same destination folder
        filesFolderName = folderName + " Files"
        parentFolder = self.current_folder()
        try:
            siblings = self.trimbleConnectClient.get_children(parentFolder)
        except Exception as ex:
            self.syncBtn.IsEnabled = True
            self.error.Content = "Job file downloaded, but couldn't check for '" + filesFolderName + "': " + str(ex)
            return

        filesFolder = next((c for c in siblings if c.IsFolder and str(c.FileName) == filesFolderName), None)

        if filesFolder is not None:
            self.statusLabel.Text = "Downloading " + filesFolderName + "..."
            self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
            try:
                self.download_connect_folder(filesFolder, os.path.join(controllerDataFolder, filesFolderName))
            except Exception as ex:
                self.syncBtn.IsEnabled = True
                self.error.Content = "Job file downloaded, but '" + filesFolderName + "' folder copy failed: " + str(ex)
                return

        # archive the job (and its Files subfolder, if any) on Trimble Connect itself, into an "old job"
        # folder alongside where they were found. get_or_create_child_folder does its own "does it already
        # exist" check via CreateFolder. Each item is archived independently (see archive_item_on_connect):
        # a real move via the Connect Web API is tried first, and only that specific item falls back to
        # upload+delete if the move doesn't work out - so a large "Files" folder that successfully moves
        # is never re-uploaded just because something else went wrong.
        oldJobFolderEnabled = bool(self.oldJobFolderEnabledCheckbox.IsChecked)
        oldJobFolderName = self.oldJobFolderBox.Text.strip()
        if oldJobFolderEnabled and oldJobFolderName:
            self.statusLabel.Text = "Archiving on Trimble Connect..."
            self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
            try:
                oldJobFolder = self.trimbleConnectClient.get_or_create_child_folder(parentFolder, oldJobFolderName)
                self.archive_item_on_connect(connectFile, False, fileName, parentFolder, oldJobFolder, downloadPath)
                if filesFolder is not None:
                    self.archive_item_on_connect(filesFolder, True, filesFolderName, parentFolder, oldJobFolder,
                                                  os.path.join(controllerDataFolder, filesFolderName))
            except Exception as ex:
                self.syncBtn.IsEnabled = True
                self.error.Content = "Downloaded locally, but archiving on Trimble Connect failed: " + str(ex)
                self.refresh_file_list()
                return

            self.refresh_file_list()

        registerError = None
        if self.jobRegisterPath:
            columnLetter = self.jobRegisterColumnBox.Text.strip().upper()
            if not columnLetter:
                registerError = "job register column is not set"
            else:
                try:
                    minimumRow = int(self.jobRegisterStartRowBox.Text.strip() or "1")
                except ValueError:
                    minimumRow = 1

                self.statusLabel.Text = "Logging to job register..."
                self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
                try:
                    ExcelJobRegister().append_value(self.jobRegisterPath, columnLetter, folderName, minimumRow)
                except Exception as ex:
                    registerError = str(ex)

        self.syncBtn.IsEnabled = True
        self.statusLabel.Text = "Downloaded to: " + downloadPath
        if filesFolder is not None:
            self.statusLabel.Text += " (with " + filesFolderName + ")"
        if oldJobFolderEnabled and oldJobFolderName:
            self.statusLabel.Text += " - archived on Connect in '" + oldJobFolderName + "'"
        if registerError is not None:
            self.error.Content = "Synced, but logging to job register failed: " + registerError
        elif self.jobRegisterPath:
            self.statusLabel.Text += " - logged to job register"

        try:
            subprocess.Popen(["explorer", targetFolder])
        except Exception:
            pass  # never let opening Explorer turn an otherwise-successful sync into an error

    def sync_field_data_job(self, job):
        # Field Data jobs live in the ConnectedWorkspace jobs-service, not Trimble Connect's own file
        # tree, so none of the Connect-specific steps apply (no "Files" sibling folder, no archiving on
        # Connect - there's nothing there to move). Each attachment already comes with its own real
        # FileName/DownloadUrl (via GetJobAttachmentDetails(jobId, includeDownloadUrl=True, ...)), so
        # downloading them straight into the destination folder reaches the same end state a
        # download-as-zip-then-extract would, without needing the zip-manifest endpoint (which 401s -
        # it needs separate Trimble Developer app/API access we don't have).
        jobName = _get_clr_property(job, "Name", None) or str(job)
        match = re.match(r"^(\d{2})(\d{2})\d{2}", jobName)
        if match is None:
            self.error.Content = "Field Data job name does not start with a 6-digit date (YYMMDD): " + jobName
            return
        yy, mm = match.group(1), match.group(2)

        if not self.localSyncFolder or not os.path.isdir(self.localSyncFolder):
            self.error.Content = "Set a valid local sync folder first."
            return

        controllerDataEnabled = bool(self.controllerDataEnabledCheckbox.IsChecked)
        tbcEnabled = bool(self.tbcEnabledCheckbox.IsChecked)
        controllerDataName = self.controllerDataBox.Text.strip()
        tbcName = self.tbcBox.Text.strip()
        if controllerDataEnabled and not controllerDataName:
            self.error.Content = "Set the 'Controller Data' subfolder name, or untick its checkbox."
            return
        if tbcEnabled and not tbcName:
            self.error.Content = "Set the 'TBC' subfolder name, or untick its checkbox."
            return

        try:
            targetFolder = os.path.join(self.localSyncFolder, *self.render_folder_structure(yy, mm, jobName))
        except Exception as ex:
            self.error.Content = str(ex)
            return
        controllerDataFolder = os.path.join(targetFolder, controllerDataName) if controllerDataEnabled else targetFolder

        try:
            foldersToCreate = [targetFolder]
            if controllerDataEnabled:
                foldersToCreate.append(controllerDataFolder)
            if tbcEnabled:
                foldersToCreate.append(os.path.join(targetFolder, tbcName))
            for folder in foldersToCreate:
                if not os.path.isdir(folder):
                    os.makedirs(folder)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        jobId = _get_clr_property(job, "JobId", None) or _get_clr_property(job, "Id", None)
        if not jobId:
            self.error.Content = "Selected Field Data job has no usable id - unexpected response shape"
            return

        self.syncBtn.IsEnabled = False
        self.statusLabel.Text = "Loading Field Data attachments for " + jobName + "..."
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

        try:
            attachments = self.connectedWorkspaceClient.get_attachments(jobId)
        except Exception as ex:
            self.syncBtn.IsEnabled = True
            self.error.Content = "Could not load Field Data attachments: " + str(ex)
            return

        if not attachments:
            self.syncBtn.IsEnabled = True
            self.error.Content = "Selected Field Data job has no attachments to download"
            return

        for a in attachments:
            fileName = _get_clr_property(a, "FileName", None)
            downloadUrl = _get_clr_property(a, "DownloadUrl", None)
            if not fileName or not downloadUrl:
                self.syncBtn.IsEnabled = True
                self.error.Content = "Attachment is missing a FileName/DownloadUrl - unexpected response shape"
                return
            self.statusLabel.Text = "Downloading " + fileName + "..."
            self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
            try:
                _download_url_to_file(downloadUrl, os.path.join(controllerDataFolder, fileName))
            except Exception as ex:
                self.syncBtn.IsEnabled = True
                self.error.Content = "Failed downloading " + fileName + ": " + str(ex)
                return

        self.syncBtn.IsEnabled = True
        self.statusLabel.Text = str(len(attachments)) + " Field Data file(s) downloaded to: " + controllerDataFolder

        # EXPERIMENTAL TEST - try flipping the job's status to CLOSED now that its download has finished,
        # via the Maps API captured behind Connect's Field Data Map Viewer (see SurveyJobsApi). Confirmed
        # working live with RESOLVED; switched to CLOSED per user request. Best-effort only and never fatal
        # to the sync itself - if it ever fails again (token scope, unmapped region, etc.) this should just
        # report the failure, not block a successful download.
        try:
            token = self.trimbleConnectClient.get_access_token()
            if not token:
                self.statusLabel.Text += " - status test skipped: no access token available"
            elif not self._fieldDataRegion or not self._fieldDataProjectId:
                self.statusLabel.Text += " - status test skipped: no region/project on record"
            else:
                SurveyJobsApi(token, self._fieldDataRegion).set_status(self._fieldDataProjectId, jobId, "CLOSED")
                self.statusLabel.Text += " - job status set to CLOSED (test)"
        except Exception as ex:
            self.statusLabel.Text += " - status test failed: " + str(ex)

        registerError = None
        if self.jobRegisterPath:
            columnLetter = self.jobRegisterColumnBox.Text.strip().upper()
            if not columnLetter:
                registerError = "job register column is not set"
            else:
                try:
                    minimumRow = int(self.jobRegisterStartRowBox.Text.strip() or "1")
                except ValueError:
                    minimumRow = 1

                self.statusLabel.Text = "Logging to job register..."
                self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
                try:
                    ExcelJobRegister().append_value(self.jobRegisterPath, columnLetter, jobName, minimumRow)
                except Exception as ex:
                    registerError = str(ex)

            self.statusLabel.Text = str(len(attachments)) + " Field Data file(s) downloaded to: " + controllerDataFolder
            if registerError is not None:
                self.error.Content = "Synced, but logging to job register failed: " + registerError
            else:
                self.statusLabel.Text += " - logged to job register"

        # EXPERIMENTAL - PARKED. Tried whether the jobs-service's own _sync action pushes this job's files
        # into the project's normal Connect folder structure, which would make the download above
        # unnecessary. job.Attributes.LinkedFiles (only populated by the per-job GetJobDetails call, not
        # by GetJobs' list) contains job TRNs of the form "trn:jobs:connect-file:{region}:{fileId}" (per
        # Trimble's public docs), but the live _sync endpoint rejected every variant tried - those exact
        # LinkedFiles values, a plain "trn:jobs:connect:{fileId}", and a couple of guessed provider names -
        # all as "Invalid job TRN format", with no example value in Swagger to go on. Left commented out
        # rather than deleted in case Trimble support (field-systems-data-service@trimble.com) or updated
        # docs clarify the right format later - re-enable by uncommenting this block.
        # if self._fieldDataProjectTrn:
        #     try:
        #         jobDetails = self.connectedWorkspaceClient.get_job_details(jobId)
        #     except Exception as ex:
        #         jobDetails = None
        #         self.statusLabel.Text += " - could not load job details for Connect sync request: " + str(ex)
        #     attributes = _get_clr_property(jobDetails, "Attributes", None)
        #     linkedFiles = list(_get_clr_property(attributes, "LinkedFiles", None) or [])
        #     if jobDetails is not None and not linkedFiles:
        #         self.statusLabel.Text += " - job has no LinkedFiles, skipped Connect sync request"
        #     for linkedFileTrn in linkedFiles:
        #         fileId = str(linkedFileTrn).rsplit(":", 1)[-1]
        #         candidateTrn = "trn:jobs:connect:" + fileId
        #         try:
        #             token = self.trimbleConnectClient.get_access_token()
        #             _sync_field_data_job_with_connect(token, self._fieldDataProjectTrn, candidateTrn)
        #             self.statusLabel.Text += " - requested Connect sync for " + candidateTrn
        #         except Exception as ex:
        #             self.statusLabel.Text += " - Connect sync request failed for " + candidateTrn + ": " + str(ex)

        try:
            subprocess.Popen(["explorer", controllerDataFolder])
        except Exception:
            pass  # never let opening Explorer turn an otherwise-successful sync into an error

    def archive_item_on_connect(self, remoteItem, isFolder, name, parentFolder, oldJobFolder, localPathForFallback):
        # tries a real move first (PATCH .../{id} with parentId, via the actual Trimble Connect Web API -
        # near-instant regardless of file size, unlike re-uploading). Confirmed via a live diagnostic dump:
        # RemoteFileId matches the item's own "ID" field exactly, and RootOriginBaseURL is the item's real
        # region-specific API host - so any failure here now would mean something else entirely (network,
        # permissions, an expired token), and just falls through to the proven-safe fallback: upload the
        # local copy, then delete the original.
        moved = False
        try:
            accessToken = self.trimbleConnectClient.get_access_token()
            itemId = getattr(remoteItem, "RemoteFileId", None)
            targetFolderId = getattr(oldJobFolder, "RemoteFileId", None)
            baseUrl = trimble_connect_region_base_url(remoteItem) or trimble_connect_region_base_url(oldJobFolder)
            if accessToken and itemId and targetFolderId and baseUrl:
                webApi = TrimbleConnectWebApi(accessToken, baseUrl)
                if isFolder:
                    webApi.move_folder(itemId, targetFolderId)
                else:
                    webApi.move_file(itemId, targetFolderId)
                moved = True
        except Exception:
            moved = False

        if moved:
            return

        if isFolder:
            oldSubFolder = self.trimbleConnectClient.get_or_create_child_folder(oldJobFolder, name)
            self.upload_connect_folder(localPathForFallback, oldSubFolder)
            self.trimbleConnectClient.remove_folder(name, parentFolder)
        else:
            self.trimbleConnectClient.save_file_remotely(localPathForFallback, oldJobFolder)
            self.trimbleConnectClient.remove_file(name, parentFolder)

    def download_connect_folder(self, remoteFolder, localFolder):
        # recursively mirrors a whole Connect folder (e.g. a job's "<name> Files" subfolder) into
        # localFolder, preserving its own subfolder structure
        if not os.path.isdir(localFolder):
            os.makedirs(localFolder)

        for child in self.trimbleConnectClient.get_children(remoteFolder):
            childLocalPath = os.path.join(localFolder, str(child.FileName))
            if child.IsFolder:
                self.download_connect_folder(child, childLocalPath)
            else:
                self.trimbleConnectClient.download_file(child, childLocalPath)
                if not self.wait_for_file(childLocalPath, expectedSize=child.Size):
                    raise Exception("Download reported success, but the file is not at: " + childLocalPath)

    def upload_connect_folder(self, localFolder, remoteFolder):
        # mirror image of download_connect_folder - re-uploads the local copy of a folder tree (e.g. the
        # "<name> Files" folder we just downloaded) into remoteFolder, recreating subfolders as needed
        for name in os.listdir(localFolder):
            fullLocalPath = os.path.join(localFolder, name)
            if os.path.isdir(fullLocalPath):
                childRemoteFolder = self.trimbleConnectClient.get_or_create_child_folder(remoteFolder, name)
                self.upload_connect_folder(fullLocalPath, childRemoteFolder)
            else:
                self.trimbleConnectClient.save_file_remotely(fullLocalPath, remoteFolder)

    def wait_for_file(self, path, expectedSize=None, timeoutSeconds=30.0, pollIntervalSeconds=0.2, stableChecksNeeded=3):
        # EndDownloadFile() returning doesn't guarantee the destination file is fully written yet - on a
        # slow connection the file can appear (even at 0 bytes) while the actual content is still being
        # streamed in, so checking os.path.isfile() alone reports a false success. When the remote node's
        # Size (IFileInformation.Size) is known, wait for the local file to actually reach it - the exact
        # check. Otherwise fall back to waiting for the size to stop growing for a few consecutive polls.
        deadline = time.time() + timeoutSeconds
        lastSize = -1
        stableChecks = 0
        while time.time() < deadline:
            if os.path.isfile(path):
                size = os.path.getsize(path)
                if expectedSize is not None and expectedSize > 0:
                    if size >= expectedSize:
                        return True
                elif size == lastSize and size > 0:
                    stableChecks += 1
                    if stableChecks >= stableChecksNeeded:
                        return True
                else:
                    stableChecks = 0
                lastSize = size
            time.sleep(pollIntervalSeconds)
        return os.path.isfile(path) and os.path.getsize(path) > 0

    def help_clicked(self, sender, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#SCR_DayJobSync")

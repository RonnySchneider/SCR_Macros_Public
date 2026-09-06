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
import webbrowser
exec(open("C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_Imports.py").read())

CIVILLO_BASE_URL = "https://app.civillo.com/api/v1"
CIVILLO_PLACEHOLDER_KEY = "PASTE_YOUR_API_KEY_HERE"
CIVILLO_PLACEHOLDER_SECRET = "PASTE_YOUR_API_SECRET_HERE"

PROPELLER_BASE_URL = "https://api.propelleraero.com/v1"
PROPELLER_PLACEHOLDER_TOKEN = "PASTE_YOUR_ACCESS_TOKEN_HERE"

JPEG_MAX_DIMENSION_PX = 65500  # libjpeg's hard per-dimension limit, independent of any file-size setting

# LAZ point cloud -> LandXML surface settings for the schedule run flow, below - now exposed as
# per-entry UI settings too (civLazMethodRadios/civLazSpacingBox in the Add Propeller Sync dialog,
# saved as civilloLazResampleMethod/civilloLazResampleSpacing per entry), so these two are only the
# entry.get(key, DEFAULT) fallback used when an entry doesn't have its own value saved (an old entry
# from before that UI existed, or an invalid spacing that failed to parse - see ok_clicked). RESAMPLE_
# SPACING is in the LAZ file's own coordinate units (typically meters); None disables resampling
# entirely and triangulates every point.
LAZ_RESAMPLE_SPACING = 1.0
LAZ_RESAMPLE_METHOD = "voxel"  # "voxel" (fast, density-target), "octree" (slower, true minimum spacing), or "grid" (regular XY grid, usually fastest for a dense uniform cloud - see resample_points_grid_2d)


class PropellerRunCancelled(Exception):
    """Raised from set_run_progress() when the user clicks Cancel mid-run - caught separately from a
    real failure so it's reported as "cancelled" rather than logged/shown as an error."""
    pass


class CivilloClient(object):
    """Thin wrapper around the Civillo REST API (https://docs.civillo.com/api/) shared by the main
    window and the Add Sync dialog, so both use the same auth/error handling."""

    # class-level (not per-instance) so the 1-request-per-second throttle holds across every
    # CivilloClient created during this session - e.g. reload_orgs_clicked makes a fresh instance
    _last_request_time = None

    def __init__(self, config):
        self.config = config

    def _throttle(self):
        # Civillo's API rejects requests faster than 1/second, so pace every real HTTP call here
        now = time.time()
        if CivilloClient._last_request_time is not None:
            elapsed = now - CivilloClient._last_request_time
            if elapsed < 1.0:
                time.sleep(1.1 - elapsed)
        CivilloClient._last_request_time = time.time()

    def _auth_header(self):
        raw = self.config["api_key"] + ":" + self.config["api_secret"]
        return "Bearer " + Convert.ToBase64String(Encoding.UTF8.GetBytes(raw))

    def get(self, relativePath):
        self._throttle()
        request = HttpWebRequest.Create(CIVILLO_BASE_URL + relativePath)
        request.Method = "GET"
        request.Headers.Add("Authorization", self._auth_header())
        return self._read_json_response(request, relativePath)

    def post_json(self, relativePath, bodyDict):
        self._throttle()
        request = HttpWebRequest.Create(CIVILLO_BASE_URL + relativePath)
        request.Method = "POST"
        request.ContentType = "application/json"
        request.Headers.Add("Authorization", self._auth_header())

        payload = Encoding.UTF8.GetBytes(json.dumps(bodyDict))
        request.ContentLength = payload.Length
        stream = request.GetRequestStream()
        stream.Write(payload, 0, payload.Length)
        stream.Close()

        return self._read_json_response(request, relativePath)

    def upload_files(self, processPath, token, application, jobId, files, progressCallback=None):
        # files: list of (localPath, uploadFileName) tuples - each becomes its own "file" part in one
        # multipart POST. The job's fileNames (from step 1) must list exactly these uploadFileNames,
        # in the same count - Civillo's processor rejects the job if the uploaded files don't match.
        self._throttle()
        primaryExt = os.path.splitext(files[0][1])[1].lstrip(".")
        url = processPath + "?token=" + token + "&application=" + application + "&jobID=" + str(jobId) + "&filetype=" + primaryExt

        boundary = "----SCRCloudReviseBoundary" + str(Guid.NewGuid()).replace("-", "")
        footer = "\r\n--" + boundary + "--\r\n"
        footerBytes = Encoding.UTF8.GetBytes(footer)

        # header bytes + each file's on-disk size (not read into memory yet - see the streaming write
        # below) is enough to compute the total Content-Length up front
        headers = []
        totalLength = 0
        for localPath, uploadFileName in files:
            header = "--" + boundary + "\r\n" + 'Content-Disposition: form-data; name="file"; filename="' + uploadFileName + '"\r\n' + "Content-Type: application/octet-stream\r\n\r\n"
            headerBytes = Encoding.UTF8.GetBytes(header)
            headers.append(headerBytes)
            totalLength += headerBytes.Length + os.path.getsize(localPath)
        totalLength += footerBytes.Length

        request = HttpWebRequest.Create(url)
        request.Method = "POST"
        request.ContentType = "multipart/form-data; boundary=" + boundary
        request.ContentLength = totalLength
        # HttpWebRequest.Timeout defaults to 100 seconds for the WHOLE request (not just connecting) -
        # fine for the small metadata GET/POST calls elsewhere in this class, but a real resampled
        # surface/orthophoto upload can be hundreds of MB (or more), which routinely takes longer than
        # that on a normal connection and gets the request aborted mid-upload ("The request was aborted:
        # The request was canceled" - confirmed against a real 205MB LandXML failing this way). The call
        # below already blocks on GetResponse() until Civillo itself sends back the upload's completion
        # response either way - that response IS the real "upload finished" signal - so there's nothing
        # to gain from capping it at some guessed duration (30 minutes still wasn't guaranteed to be
        # enough for a multi-GB file over a slow/mobile connection); Infinite just means this is bounded
        # by Civillo actually responding (or a genuine network failure) instead of an arbitrary ceiling.
        request.Timeout = -1           # System.Threading.Timeout.Infinite - documented by Microsoft as
        request.ReadWriteTimeout = -1  # the "no timeout" value for both of these specific properties
        # AllowWriteStreamBuffering defaults to True, which makes HttpWebRequest buffer everything written
        # to the request stream in memory FIRST and only actually push it out over the network once the
        # request is finalized (GetResponse()/the stream closing) - for a large upload that looks exactly
        # like nothing is happening (no outgoing network traffic, Civillo's own UI stuck on "waiting for
        # upload") while this loop is still just filling RAM. False makes each Write() go straight to the
        # socket instead, which is also required for the chunked-read loop below to have any real effect.
        request.AllowWriteStreamBuffering = False

        # stream each file straight from disk into the request body instead of loading it whole into a
        # byte[] first (the previous File.ReadAllBytes(localPath) approach) - .NET arrays are indexed by
        # Int32, so any file >= 2GB threw "The file is too long. This operation is currently limited to
        # supporting files less than 2 gigabytes in size." outright, confirmed against a real oversized
        # orthophoto TIFF upload. Chunked streaming has no such limit and holds only one small buffer in
        # memory regardless of file size.
        bytesSent = 0
        with request.GetRequestStream() as stream:
            buffer = Array.CreateInstance(Byte, 65536)
            for (localPath, uploadFileName), headerBytes in zip(files, headers):
                stream.Write(headerBytes, 0, headerBytes.Length)
                bytesSent += headerBytes.Length
                with FileStream(localPath, FileMode.Open, FileAccess.Read) as inStream:
                    while True:
                        bytesRead = inStream.Read(buffer, 0, buffer.Length)
                        if bytesRead <= 0:
                            break
                        stream.Write(buffer, 0, bytesRead)
                        bytesSent += bytesRead
                        # AllowWriteStreamBuffering=False (above) means this Write() really did just push
                        # bytesRead further bytes onto the socket, so bytesSent/totalLength here is a real
                        # upload progress fraction, not a "how much have I handed to a buffer" number
                        if progressCallback is not None:
                            progressCallback(bytesSent, totalLength)
            stream.Write(footerBytes, 0, footerBytes.Length)
            bytesSent += footerBytes.Length

        return self._read_json_response(request, processPath, allowNonJson=True)

    def _read_json_response(self, request, relativePath, allowNonJson=False):
        try:
            response = request.GetResponse()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("Civillo API call to " + relativePath + " failed: " + str(ex.Message) + " " + detail)

        with StreamReader(response.GetResponseStream(), Encoding.UTF8) as sr:
            body = sr.ReadToEnd()

        if body == "":
            return {}
        try:
            return json.loads(body)
        except Exception:
            if allowNonJson:
                return {"raw": body}
            raise


class PropellerClient(object):
    """Thin wrapper around the Propeller public API (https://propelleraero.readme.io/), read-only for
    now - lists organizations/sites/surveys and their downloadable files so a future sync flow can pull
    a Propeller survey file the same way CivilloClient pushes one."""

    def __init__(self, config):
        self.config = config

    def _auth_headers(self, request):
        token = self.config["access_token"]
        if token.lower().startswith("bearer "):
            token = token[len("bearer "):]
        request.Headers.Add("Authorization", "Bearer " + token)
        request.Accept = "application/json"

    def _build_url(self, relativePath, params):
        query = "&".join(Uri.EscapeDataString(str(k)) + "=" + Uri.EscapeDataString(str(v)) for k, v in params.items())
        return PROPELLER_BASE_URL + relativePath + "?" + query

    def _get_with_retry(self, relativePath, params):
        # a freshly-created/rotated Propeller access token can 403 with an IAM-style "ACCESSDENIED ...
        # explicit deny in an identity-based policy" for the first several seconds after creation, before
        # the token's permissions finish propagating on Propeller's side - this is transient, so retry a
        # few times with a short backoff rather than surfacing a scary permissions error immediately
        maxAttempts = 5
        delaySeconds = 1.5
        lastEx = None

        for attempt in range(maxAttempts):
            request = HttpWebRequest.Create(self._build_url(relativePath, params))
            request.Method = "GET"
            self._auth_headers(request)

            try:
                return request.GetResponse()
            except WebException as ex:
                detail = ""
                if ex.Response is not None:
                    with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                        detail = sr.ReadToEnd()
                lastEx = Exception("Propeller API call to " + relativePath + " failed: " + str(ex.Message) + " " + detail)

                if "ACCESSDENIED" not in detail or attempt == maxAttempts - 1:
                    raise lastEx
                time.sleep(delaySeconds)

        raise lastEx

    def get_list(self, relativePath, extraParams=None):
        # Propeller paginates with page/limit query params and a {"results": [...], "pagination": {"count", "limit", "page"}}
        # response shape (confirmed against https://propelleraero.readme.io/reference/list_*.md) - keep requesting
        # pages at the max page size (500) until we've collected "count" items.
        allItems = []
        page = 1  # docs say page is 0-indexed (minimum 0), but the live API 404s ("Page out of range") on page 0
        while True:
            params = dict(extraParams or {})
            params["page"] = page
            params["limit"] = 500

            response = self._get_with_retry(relativePath, params)

            with StreamReader(response.GetResponseStream(), Encoding.UTF8) as sr:
                body = sr.ReadToEnd()

            parsed = json.loads(body) if body != "" else {}
            results = parsed.get("results", [])
            allItems.extend(results)

            pagination = parsed.get("pagination", {})
            if len(allItems) >= pagination.get("count", len(allItems)) or len(results) == 0:
                break
            page += 1

        return allItems

    def get_organizations(self):
        return self.get_list("/organizations")

    def get_sites(self, organizationId):
        return self.get_list("/organizations/" + str(organizationId) + "/sites")

    def get_surveys(self, organizationId, siteId):
        # date_captured_gt/lt are documented as optional but the server 400s without at least one -
        # pass a range wide enough to cover everything ever captured through tomorrow (clock-skew margin)
        dateRangeParams = {
            "date_captured_gt": "1970-01-01T00:00:00Z",
            "date_captured_lt": (datetime.utcnow() + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        return self.get_list("/organizations/" + str(organizationId) + "/sites/" + str(siteId) + "/surveys", dateRangeParams)

    def get_survey_files(self, organizationId, siteId, surveyId):
        return self.get_list("/organizations/" + str(organizationId) + "/sites/" + str(siteId) + "/surveys/" + str(surveyId) + "/files")

    def download_file(self, url, localPath, progressCallback=None):
        # downloads to "<localPath>.part" first, resuming from wherever a previous attempt left off (via
        # an HTTP Range request) rather than restarting from scratch - large orthophoto deliverables can
        # be multiple GB, so an interrupted download shouldn't mean starting over. Only renames to the
        # final localPath once the whole file has actually been received.
        partPath = localPath + ".part"
        resumeOffset = os.path.getsize(partPath) if os.path.isfile(partPath) else 0

        request = HttpWebRequest.Create(url)
        request.Method = "GET"
        # survey file URLs are typically pre-signed (e.g. S3) and don't need our bearer token - only
        # attach it when the URL actually points back at Propeller's own API
        if "propelleraero.com" in url:
            self._auth_headers(request)
        if resumeOffset > 0:
            request.AddRange(resumeOffset)

        try:
            response = request.GetResponse()
        except WebException as ex:
            detail = ""
            if ex.Response is not None:
                with StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8) as sr:
                    detail = sr.ReadToEnd()
            raise Exception("Propeller file download failed: " + str(ex.Message) + " " + detail)

        # a 206 confirms the server actually honored our Range request; if we asked to resume but got a
        # plain 200 back, the server ignored it (or doesn't support ranges for this URL) - appending to
        # the partial file in that case would produce a corrupt result, so start over instead
        resumed = resumeOffset > 0 and int(response.StatusCode) == 206
        fileMode = FileMode.Append if resumed else FileMode.Create
        bytesWritten = resumeOffset if resumed else 0
        totalBytes = response.ContentLength + bytesWritten if resumed else response.ContentLength

        with response.GetResponseStream() as inStream:
            with FileStream(partPath, fileMode) as outStream:
                buffer = Array.CreateInstance(Byte, 65536)
                while True:
                    bytesRead = inStream.Read(buffer, 0, buffer.Length)
                    if bytesRead <= 0:
                        break
                    outStream.Write(buffer, 0, bytesRead)
                    bytesWritten += bytesRead
                    if progressCallback is not None:
                        progressCallback(bytesWritten, totalBytes)

        if os.path.isfile(localPath):
            os.remove(localPath)
        os.rename(partPath, localPath)


_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def sanitize_filename(name):
    # "." is also replaced (not just Windows' invalid characters) because Civillo's server-side upload
    # validator appears to split a filename's extension on the *first* dot rather than the last - a
    # survey name containing a literal period (e.g. "...[State 280.000]...") produced an uploaded filename
    # whose "extension" it parsed as everything after that dot, rejecting it as invalid
    replaced = _INVALID_FILENAME_CHARS + "."
    return "".join(c if c not in replaced else "_" for c in name)


def derive_output_filename(surveyName, pixelSizeMeters, ext):
    # <survey name>_<pixel size in cm>cm.<ext> - e.g. "022305 CDSA Survey_10cm.tif" for a 0.1m/pixel TIFF.
    # Pixel size in cm (not raw meters) keeps the number short and reads naturally for typical GSDs.
    pixelSizeCm = int(round(pixelSizeMeters * 100))
    return sanitize_filename(surveyName) + "_" + str(pixelSizeCm) + "cm." + ext.lstrip(".")


def worldfile_path_for(outputPath):
    # GDAL's WORLDFILE=YES creation option always writes ".wld" regardless of image format - unlike the
    # driver-specific ESRI convention (.tfw/.jgw) some other tools use, confirmed against actual output
    base, _ = os.path.splitext(outputPath)
    return base + ".wld"



# Civillo rejects GDAL's generic ".wld" as an "invalid extension" - it expects a format-specific
# convention instead (.tfw for GeoTIFF, .jgw for JPEG - the standard ESRI worldfile convention).
# Civillo's own docs literally say ".jpw" instead of ".jgw" - tried that once, confirmed wrong: it got
# the exact same "invalid extensions" 400 error .wld does, so .jpw isn't accepted either; .jgw is the
# one that doesn't get rejected outright. A real upload with .jgw still got rejected downstream for
# "exceeding Civillo's 35km x 35km limit" on an image whose real extent (per its own worldfile) was
# only ~1.25km x 1km - so .jgw being *accepted* doesn't yet prove Civillo actually reads and applies
# its pixel size; that root cause is still open (see conversation), just not "wrong extension".
WORLDFILE_EXTENSION_MAP = {"tif": "tfw", "tiff": "tfw", "jpg": "jgw", "jpeg": "jgw"}


def final_worldfile_path_for(outputPath, ext):
    # pure path computation (no I/O) - lets a cache-hit check work out where the worldfile WOULD be
    # without renaming anything, mirroring what rename_worldfile_for_format below actually produces
    targetExt = WORLDFILE_EXTENSION_MAP.get(ext.lower())
    if targetExt is None:
        return None
    return os.path.splitext(outputPath)[0] + "." + targetExt


def rename_worldfile_for_format(worldfilePath, ext):
    targetExt = WORLDFILE_EXTENSION_MAP.get(ext.lower())
    if targetExt is None:
        return worldfilePath

    renamedPath = os.path.splitext(worldfilePath)[0] + "." + targetExt
    if os.path.isfile(renamedPath):
        os.remove(renamedPath)
    os.rename(worldfilePath, renamedPath)
    return renamedPath


class PropellerDownloadCache(object):
    """Manages the local Propeller-Cache download folder. Propeller's survey deliverables (orthomosaic,
    DTM/DSM, pointcloud) are generated once and never change, so a file already present here never needs
    re-downloading - the local path (a <site name>/<survey name> subfolder + sanitized file name) is the
    cache key. Files (and, transitively, empty subfolders) untouched for MAX_AGE_DAYS are purged so the
    folder doesn't grow forever."""

    MAX_AGE_DAYS = 30

    def __init__(self, folder):
        self.folder = folder

    def local_path(self, siteName, surveyName, fileInfo):
        surveyFolder = os.path.join(self.folder, sanitize_filename(siteName), sanitize_filename(surveyName))
        base = sanitize_filename(fileInfo.get("name", "file"))
        ext = "." + fileInfo.get("format", "bin")
        if not base.lower().endswith(ext.lower()):
            base += ext
        return os.path.join(surveyFolder, base)

    def get_or_download(self, propellerClient, siteName, surveyName, fileInfo, progressCallback=None):
        # returns (path, wasCached) - wasCached lets the caller show a "found in cache" note instead of
        # silently skipping straight past what would otherwise look like a missing download step
        path = self.local_path(siteName, surveyName, fileInfo)
        folder = os.path.dirname(path)
        if not os.path.isdir(folder):
            os.makedirs(folder)

        if os.path.isfile(path):
            os.utime(path, None)  # touch - keep files still in use from aging out of cleanup()
            return path, True

        propellerClient.download_file(fileInfo["url"], path, progressCallback)
        return path, False

    def cleanup(self):
        if not os.path.isdir(self.folder):
            return
        cutoff = time.time() - self.MAX_AGE_DAYS * 24 * 3600
        for root, dirs, files in os.walk(self.folder, topdown=False):
            for name in files:
                full = os.path.join(root, name)
                if os.path.getmtime(full) < cutoff:
                    os.remove(full)
            if root != self.folder and not os.listdir(root):
                os.rmdir(root)


def zip_files_for_civillo_ortho(files):
    """Bundles a resampled orthophoto with its worldfile into a single zip before handing it to
    push_to_civillo/push_to_civillo_create_new - confirmed by hand (real test layer, real Civillo
    project) that uploading them as two separate multipart files succeeds at the HTTP level
    ("success": true) but Civillo's raster processor silently ignores the worldfile and defaults to a
    nominal ~1 m/pixel scale instead, which then gets rejected downstream for "exceeding" its 35km x
    35km processing limit on an image whose real extent (per the ignored worldfile) was nowhere close.
    Zipping them together fixed it in that same test. Every existing image layer in that live project
    was itself sourced from a ".zip" too, confirming this is simply the format Civillo's raster upload
    endpoint (a different backend than the vector/CAD one below) actually expects.

    Only meaningful when there's a worldfile to bundle (len(files) > 1) - an unmodified GeoTIFF
    passthrough has its georeferencing embedded already and is uploaded as-is. NOT used for the
    LAZ->LandXML path (self-contained, no worldfile) or the 12d/linestyle Local->Civillo path in
    run_single_sync (a different Civillo endpoint that does NOT unzip - confirmed separately, zipping
    there was tried and abandoned - see [[project_scr_transferhub]])."""
    if len(files) < 2:
        return files

    imagePath, imageName = files[0]
    zipPath = os.path.splitext(imagePath)[0] + ".zip"
    zipName = os.path.splitext(imageName)[0] + ".zip"
    # ZIP_STORED (no compression) - the image is already JPEG/TIFF-compressed and the worldfile is a few
    # bytes of text, so DEFLATE would just spend CPU time (and, for a multi-GB TIFF, a noticeable delay)
    # for a few % smaller at best; the zip here exists purely to bundle the pair together for Civillo,
    # not to shrink anything
    with zipfile.ZipFile(zipPath, "w", zipfile.ZIP_STORED) as zf:
        for path, name in files:
            zf.write(path, arcname=name)
    return [(zipPath, zipName)]


def push_to_civillo(civilloClient, orgNickname, projectId, replaceLayerId, files, progressCallback=None):
    # files: list of (localPath, uploadFileName) tuples - a single LandXML, a raw GeoTIFF passthrough
    # (embedded georeferencing is enough), or a single zip already bundling image+worldfile (see
    # zip_files_for_civillo_ortho - the caller zips a resampled orthophoto+worldfile before this point,
    # this function itself doesn't care whether "files" holds one entry or several)
    projectInfo = civilloClient.get("/" + orgNickname + "/projects/" + str(projectId))
    srid = projectInfo.get("defaultProjection", -1)

    body = {
        "mode": 1,  # 1 = revise an existing layer
        "fileNames": [name for _, name in files],
        "replaceLayerId": replaceLayerId,
        "fileLastModifieds": [int(os.path.getmtime(path)) for path, _ in files],
        "srid": srid,
    }
    initResp = civilloClient.post_json("/" + orgNickname + "/projects/" + str(projectId) + "/layers", body)
    civilloClient.upload_files(initResp["processPath"], initResp["token"], initResp["application"], initResp["job"], files, progressCallback)


def push_to_civillo_create_new(civilloClient, orgNickname, projectId, title, files, progressCallback=None):
    """Creates a brand-new Civillo layer (mode 0) named `title`, rather than revising an existing one -
    see push_to_civillo above for the revise (mode 1) equivalent this mirrors. Backs a schedule entry's
    "create additional new layer" option, mainly meant for layer types (terrain/DTM/surface) that never
    show up as a pick in the file list to revise in the first place - confirmed by hand that
    /layer-directory never returns them, only the flat /layers endpoint does (see
    get_civillo_unassigned_layers), and even a fresh create's own response never hands back the new
    layerId either, so there is nothing to remember/revise against on a later run - every run creates a
    genuinely new layer.

    Fails with Civillo's own HTTP 400 if a layer with this exact title already exists in the project
    (confirmed by hand) - not handled here, since retrying under a different title is a decision this
    function shouldn't make for the caller; a stable {YYMMDD}-based title (see format_survey_yymmdd)
    keeps re-running the same survey's sync from hitting this by accident."""
    projectInfo = civilloClient.get("/" + orgNickname + "/projects/" + str(projectId))
    srid = projectInfo.get("defaultProjection", -1)

    body = {
        "mode": 0,  # 0 = create a new layer
        "fileNames": [name for _, name in files],
        "title": title,
        "fileLastModifieds": [int(os.path.getmtime(path)) for path, _ in files],
        "srid": srid,
    }
    initResp = civilloClient.post_json("/" + orgNickname + "/projects/" + str(projectId) + "/layers", body)
    civilloClient.upload_files(initResp["processPath"], initResp["token"], initResp["application"], initResp["job"], files, progressCallback)


def format_survey_yymmdd(dateCapturedIso):
    """Formats a Propeller survey's own date_captured (ISO 8601, e.g. "2026-09-04T06:29:32Z") as
    "YYMMDD" for the {YYMMDD} substitution in a "create additional new layer" title template - matches
    the date-prefix convention already used by hand for terrain layers in this project (e.g.
    "260904 - Terrain - Dam to X2..."), and is stable across re-running the same survey's sync on a
    different calendar day (unlike using today's date, which would create a new layer every time).
    Returns None if dateCaptured is missing/unparseable, so {YYMMDD} is left untouched rather than
    silently substituting something wrong."""
    if not dateCapturedIso:
        return None
    try:
        # only the leading "YYYY-MM-DD" is ever needed - slicing it off and parsing just that avoids
        # depending on exactly how many fractional-second digits/timezone suffix the API happens to send
        dt = datetime.strptime(dateCapturedIso[:10], "%Y-%m-%d")
        return dt.strftime("%y%m%d")
    except Exception:
        return None


class TrimbleConnectClient(object):
    """Thin wrapper around Trimble.Vce.Services.Construction.TrimbleConnect.TrimbleConnectService,
    reusing TBC's own already-authenticated Trimble ID session - no separate OAuth flow needed. GetFileList(parent, progressBar) is the one
    primitive the service exposes for browsing: called with None it returns the region list; called with
    a region it returns the projects in that region; called with a project or folder it returns that
    folder's immediate children (one level at a time - there is no single call that returns a whole tree,
    unlike Civillo's layer-directory, so deep folders must be browsed into rather than eagerly flattened)."""

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
        # fall back to no progress bar if it turns out not to satisfy the interface. Unlike CivilloClient's
        # own upload_files, this call is a single opaque SaveFileRemotely from Trimble's own SDK - there's
        # no hook into its internal transport to get real byte-level progress/ETA out of, so the best this
        # can do is set a Title on TBC's native status-bar progress control so it's not just blank/silent
        # while a large upload is in flight (SaveFileRemotely itself presumably drives SetProgress/percent
        # internally, same as it would for any other caller passing it a progress bar).
        try:
            ProgressBar.TBC_ProgressBar.Title = "Uploading " + os.path.basename(localFile) + " to Trimble Connect..."
            self._get_service().SaveFileRemotely(localFile, parentFolderOrProject, ProgressBar.TBC_ProgressBar)
        except Exception:
            self._get_service().SaveFileRemotely(localFile, parentFolderOrProject, None)
        finally:
            ProgressBar.TBC_ProgressBar.Title = ""


def _describe_dotnet_exception(ex):
    # unwraps the .NET InnerException chain - a TypeInitializationException (like GdalPINVOKE's static
    # constructor failing) hides the actually useful error (e.g. DllNotFoundException/BadImageFormatException
    # naming the missing/mismatched native DLL) one level down, which str(ex) alone won't show
    netEx = getattr(ex, "clsException", ex)
    parts = [str(ex)]
    inner = getattr(netEx, "InnerException", None)
    depth = 0
    while inner is not None and depth < 5:
        parts.append("inner[" + str(depth) + "]: " + inner.GetType().FullName + ": " + str(inner.Message))
        inner = inner.InnerException
        depth += 1
    return " | ".join(parts)


def _ensure_gdal_native_dir_on_path():
    # gdal_wrap.dll (the native P/Invoke target behind gdal_csharp.dll) lives in "<TBC install>\gdal\x64\",
    # a subfolder Windows' default DLL search order does not look in (only the exe's own folder, System32,
    # and PATH are searched, not arbitrary subfolders) - this is the actual cause of the earlier
    # "Unable to load DLL 'gdal_wrap'" failure, confirmed by the file existing identically across TBC
    # 2024.01/2025.21/2026.1 installs (so not a version-specific naming mismatch). Prepending the folder
    # to PATH before GDAL is first touched lets LoadLibrary find it.
    import System.Diagnostics
    exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName
    gdalNativeDir = os.path.join(os.path.dirname(exePath), "gdal", "x64")
    if os.path.isdir(gdalNativeDir) and gdalNativeDir not in os.environ.get("PATH", ""):
        os.environ["PATH"] = gdalNativeDir + os.pathsep + os.environ.get("PATH", "")


def _log_gdal_dataset_info(dataset, path):
    # dumps what GDAL itself actually read for georeferencing (driver, size, geotransform, projection) -
    # answers directly whether GDAL is failing to find/parse the embedded georeferencing at all, versus
    # something more specific to the translate call itself
    lines = ["--- GDAL dataset info for " + path + " ---"]
    try:
        lines.append("Driver: " + dataset.GetDriver().ShortName + " / " + dataset.GetDriver().LongName)
    except Exception as ex:
        lines.append("Driver: ERROR " + str(ex))

    try:
        lines.append("Size: " + str(dataset.RasterXSize) + " x " + str(dataset.RasterYSize) + ", bands=" + str(dataset.RasterCount))
    except Exception as ex:
        lines.append("Size: ERROR " + str(ex))

    try:
        geoTransform = Array[float]([0.0, 0.0, 0.0, 0.0, 0.0, 0.0])
        dataset.GetGeoTransform(geoTransform)
        lines.append("GeoTransform: " + str(list(geoTransform)))
    except Exception as ex:
        lines.append("GeoTransform: ERROR " + str(ex))

    try:
        lines.append("ProjectionRef: " + str(dataset.GetProjectionRef()))
    except Exception as ex:
        lines.append("ProjectionRef: ERROR " + str(ex))

    try:
        if not os.path.isdir(r"C:\temp"):
            os.makedirs(r"C:\temp")
        with open(r"C:\temp\gdal_debug.log", "a") as f:
            f.write(str(datetime.now()) + "\n" + "\n".join(lines) + "\n\n")
    except Exception:
        pass


def get_geotiff_extent_meters(path):
    """Returns (extentXMeters, extentYMeters) - the real-world ground size of a GeoTIFF. Used to work out
    the coarsest pixel size a hard per-format limit (e.g. JPEG's 65500px-per-dimension cap) allows,
    before ever attempting a translate that would hit it and fail with a raw libjpeg error."""
    _ensure_gdal_native_dir_on_path()
    clr.AddReference("gdal_csharp")
    import OSGeo.GDAL as osgeo_gdal

    try:
        osgeo_gdal.Gdal.AllRegister()
    except Exception as ex:
        raise Exception("Gdal.AllRegister() failed: " + _describe_dotnet_exception(ex))
    osgeo_gdal.Gdal.SetConfigOption("GDAL_PAM_ENABLED", "NO")  # don't write .aux.xml sidecar files - unused here

    dataset = osgeo_gdal.Gdal.Open(path, osgeo_gdal.Access.GA_ReadOnly)
    if dataset is None:
        raise Exception("Could not open input file: " + path)

    try:
        geoTransform = Array[float]([0.0, 0.0, 0.0, 0.0, 0.0, 0.0])
        dataset.GetGeoTransform(geoTransform)
        return dataset.RasterXSize * abs(geoTransform[1]), dataset.RasterYSize * abs(geoTransform[5])
    finally:
        dataset.Dispose()


def get_geotiff_pixel_size_meters(path):
    """Returns the source GeoTIFF's own native ground sample distance (meters/pixel, averaged across X/Y
    in case they differ slightly) - used as the starting point for the max-GB retry loop when no pixel
    size was configured at all (see _prepare_platform_file), since there's no requested value to grow
    from otherwise."""
    _ensure_gdal_native_dir_on_path()
    clr.AddReference("gdal_csharp")
    import OSGeo.GDAL as osgeo_gdal

    try:
        osgeo_gdal.Gdal.AllRegister()
    except Exception as ex:
        raise Exception("Gdal.AllRegister() failed: " + _describe_dotnet_exception(ex))
    osgeo_gdal.Gdal.SetConfigOption("GDAL_PAM_ENABLED", "NO")

    dataset = osgeo_gdal.Gdal.Open(path, osgeo_gdal.Access.GA_ReadOnly)
    if dataset is None:
        raise Exception("Could not open input file: " + path)

    try:
        geoTransform = Array[float]([0.0, 0.0, 0.0, 0.0, 0.0, 0.0])
        dataset.GetGeoTransform(geoTransform)
        return (abs(geoTransform[1]) + abs(geoTransform[5])) / 2.0
    finally:
        dataset.Dispose()


def convert_geotiff_resolution(inputPath, outputPath, pixelSizeMeters, outputFormat, progressCallback=None):
    """Resamples a GeoTIFF to the given pixel size (in meters/pixel - i.e. ground sample distance) and
    writes it out as GTiff or JPEG plus a worldfile, using OSGeo.GDAL directly (Trimble.Vce.GE.Imaging.Gdal
    is only a thin wrapper around it and has no resample/reproject function of its own - see conversation
    notes). progressCallback, if given, is called with an int 0-100 as GDAL reports progress - this
    function has no UI of its own, so the caller decides where that goes."""
    # go through the namespace module + attribute access rather than "from X import Y" - not an IronPython
    # limitation (direct import works fine for real names, e.g. "from OSGeo.GDAL import Dataset"), just
    # convenient here since several names we reference (GDALTranslateOptions, wrapper_GDALTranslate) are
    # non-obvious in this GDAL 3.0.5 build and easier to probe/reflect via the module object
    _ensure_gdal_native_dir_on_path()
    clr.AddReference("gdal_csharp")
    import OSGeo.GDAL as osgeo_gdal

    try:
        osgeo_gdal.Gdal.AllRegister()
    except Exception as ex:
        raise Exception("Gdal.AllRegister() failed: " + _describe_dotnet_exception(ex))
    osgeo_gdal.Gdal.SetConfigOption("GDAL_PAM_ENABLED", "NO")  # don't write a .aux.xml sidecar for the output - we don't use it

    driverName = "JPEG" if outputFormat.upper() == "JPEG" else "GTiff"

    srcDataset = osgeo_gdal.Gdal.Open(inputPath, osgeo_gdal.Access.GA_ReadOnly)
    if srcDataset is None:
        raise Exception("Could not open input file: " + inputPath)

    _log_gdal_dataset_info(srcDataset, inputPath)

    # GDAL's -tr handling strictly rejects any non-zero rotation term in the geotransform, but real-world
    # exports (like this one) can carry ~1e-14 floating-point noise there instead of a true zero - many
    # orders of magnitude smaller than the pixel size, and not actually rotation. Clean it via a lightweight
    # VRT that references the source pixels (no multi-GB data duplication) with the noise zeroed out.
    geoTransform = Array[float]([0.0, 0.0, 0.0, 0.0, 0.0, 0.0])
    srcDataset.GetGeoTransform(geoTransform)
    rotationEpsilon = 1e-8
    translateSource = srcDataset
    vrtPath = None

    if (0 < abs(geoTransform[2]) < rotationEpsilon) or (0 < abs(geoTransform[4]) < rotationEpsilon):
        vrtPath = outputPath + ".source.vrt"
        vrtDataset = osgeo_gdal.Gdal.GetDriverByName("VRT").CreateCopy(vrtPath, srcDataset, 0, None, None, None)
        cleanedTransform = Array[float]([geoTransform[0], geoTransform[1], 0.0, geoTransform[3], 0.0, geoTransform[5]])
        vrtDataset.SetGeoTransform(cleanedTransform)
        vrtDataset.FlushCache()
        translateSource = vrtDataset

    # this bundled GDAL 3.0.5 build's C# bindings prefix everything with "GDAL": the options class is
    # GDALTranslateOptions (not TranslateOptions) and the method is wrapper_GDALTranslate (not Translate) -
    # confirmed by reflecting the assembly's actual type/method lists
    translateArgList = ["-tr", str(pixelSizeMeters), str(pixelSizeMeters), "-r", "bilinear", "-of", driverName, "-co", "WORLDFILE=YES"]

    if driverName == "JPEG":
        # GDAL's JPEG driver defaults QUALITY to 75 - 85 gives noticeably cleaner output for a modest
        # increase in file size (and roughly proportional encode time, since there's more detail to
        # entropy-code - see the max-size retry loop for what that does to the size-limit estimate)
        translateArgList += ["-co", "QUALITY=85"]

        if srcDataset.RasterCount > 3:
            # JPEG has no alpha channel - a 4(+)-band source (RGB + alpha, or +NIR) passed straight through
            # doesn't error, but libjpeg treats the leftover band as a 4th color component and writes an Adobe
            # CMYK-style JPEG, which most viewers (Windows Photos included) don't recognize and render with a
            # brownish/sepia cast instead of the intended colors. Restrict to the first 3 bands - orthophoto
            # exports are always band order R,G,B(,alpha/NIR...) - so the output is a plain 3-band RGB JPEG.
            translateArgList += ["-b", "1", "-b", "2", "-b", "3"]

    translateArgs = Array[str](translateArgList)
    options = osgeo_gdal.GDALTranslateOptions(translateArgs)

    # best-effort progress reporting - GDAL's progress delegate signature isn't confirmed for this
    # specific 3.0.5 build, so fall back to no percentage reporting if constructing or using the
    # delegate fails for any reason, rather than blocking the actual conversion on it
    gdalProgressDelegate = None
    # a callback exception (e.g. set_run_progress()'s PropellerRunCancelled when the user hits Cancel
    # mid-resample) is stashed here rather than let to escape _report_progress directly - it's invoked BY
    # GDAL's native code via a P/Invoke callback, and a managed exception escaping across that boundary
    # risks crashing the whole process rather than failing gracefully. Returning 0 instead tells GDAL to
    # abort the translate through its own normal (safe) cancellation path; the stashed exception is
    # re-raised below, once back in pure managed code on this side of that boundary.
    progressCallbackException = [None]
    if progressCallback is not None:
        try:
            def _report_progress(complete, message, data):
                try:
                    progressCallback(int(complete * 100))
                    return 1
                except Exception as progressEx:
                    progressCallbackException[0] = progressEx
                    return 0
            gdalProgressDelegate = osgeo_gdal.Gdal.GDALProgressFuncDelegate(_report_progress)
        except Exception:
            gdalProgressDelegate = None

    try:
        outDataset = osgeo_gdal.Gdal.wrapper_GDALTranslate(outputPath, translateSource, options, gdalProgressDelegate, None)
    except Exception:
        if progressCallbackException[0] is not None:
            outDataset = None  # aborted via the callback above - don't mask it by retrying without one
        else:
            outDataset = osgeo_gdal.Gdal.wrapper_GDALTranslate(outputPath, translateSource, options, None, None)

    if translateSource is not srcDataset:
        translateSource.Dispose()
    srcDataset.Dispose()

    if vrtPath is not None:
        try:
            os.remove(vrtPath)
        except Exception:
            pass

    if progressCallbackException[0] is not None:
        raise progressCallbackException[0]

    if outDataset is None:
        raise Exception("GDAL Translate failed for " + inputPath + " -> " + outputPath)
    outDataset.Dispose()


def parse_point_count_from_filename(path):
    """Parses a point count baked into a LAZ filename - thousands grouped with underscores instead of
    commas (commas aren't filename-safe), immediately before "_points" and the extension - e.g. both a
    bare test file like "46_403_028_points.laz" and a real Propeller deliverable name like
    "022305_CDSIP_GDA2020_..._46_403_028_points.laz" match; only the trailing digit-group run right
    before "_points.la[sz]" is captured, whatever precedes it is ignored. Returns None if the filename
    doesn't end that way rather than guessing - callers should treat this purely as a speed hint (see
    SCRLasZip.read_points' expectedCount), never as authoritative over what the file actually contains."""
    base = os.path.basename(path)
    # a coordinate zone number (e.g. "..._MGA_zone_56_470_021_points.laz") sits directly before the real
    # count in some real Propeller filenames and is otherwise indistinguishable from it by digit-grouping
    # width alone - "56_470_021" reads as a perfectly valid 3-group count on its own (56,470,021), exactly
    # as valid-looking as "470_021" (470,021) with "56" excluded, so requiring each group after the first
    # to be exactly 3 digits (ruling out e.g. "_46_403_028" where "46" breaks the chain) isn't enough by
    # itself once the zone number happens to be a width that keeps the chain intact. Recognizing the
    # "zone_NN_" marker itself and cutting everything up to and including it removes the ambiguity instead
    # of guessing from digit widths.
    zoneMatch = re.search(r"zone_\d{1,2}_", base, re.IGNORECASE)
    if zoneMatch:
        base = base[zoneMatch.end():]
    match = re.search(r"(\d{1,3}(?:_\d{3})*)_points\.la[sz]$", base, re.IGNORECASE)
    if not match:
        return None
    digits = match.group(1).replace("_", "")
    return int(digits) if digits.isdigit() else None


def extract_filename_from_propeller_url(url):
    """Propeller's file download URLs are pre-signed S3-style links that echo the real filename back via
    a response-content-disposition query parameter, in the same "attachment;filename=NAME" shape a real
    download's Content-Disposition response header would have - readable straight out of the URL string,
    no request needed. This is the filename shown in Propeller's own web UI (point count included, when
    Propeller names it that way) - notably NOT the same as the "name" field the files API itself returns
    (a shorter descriptive label like "Ortho TIFF (GDA2020 / MGA zone 56)", no point count).
    Returns None if the URL has no such parameter."""
    match = re.search(r"[?&]response-content-disposition=([^&]+)", url, re.IGNORECASE)
    if not match:
        return None
    disposition = Uri.UnescapeDataString(match.group(1))
    filenameMatch = re.search(r"filename=([^;]+)", disposition, re.IGNORECASE)
    return filenameMatch.group(1) if filenameMatch else None


def format_site_crs_label(site):
    """Formats a site's own coordinate_reference_system_pointer (an {authority, id, unit} pair - "EPSG"
    or "PROP-LOCAL" per Propeller's API) as a plain "<authority>:<id>" label, e.g. "EPSG:7856". Used as
    the file list's CRS group label instead of trying to parse a human-readable name out of a filename -
    that only worked for Australia's "zone_NN" naming convention (see conversation history); an EPSG/
    PROP-LOCAL code straight from the site itself is unambiguous and correct everywhere in the world,
    at the cost of showing a code instead of a friendly name like "GDA2020 / MGA zone 56".
    Returns None if the site has no usable pointer."""
    pointer = (site or {}).get("coordinate_reference_system_pointer") or {}
    horizontal = pointer.get("horizontal") or {}
    authority = horizontal.get("authority")
    crsId = horizontal.get("id")
    if not authority or not crsId:
        return None
    return str(authority) + ":" + str(crsId)


def _run_pdal_pipeline(pdalExePath, pipelinePath, statusText, progressCallback=None):
    """Runs `pdal.exe pipeline <pipelinePath>` as a subprocess and blocks until it exits - shared by both
    pdal_resample_laz (thin only) and _pdal_resample_and_triangulate_to_landxml (thin + Delaunay in one
    pipeline). Shells out rather than P/Invoking into PDAL's own DLL: pdalcpp.dll exports C++ classes
    (name-mangled, compiler/ABI-specific), not a stable, P/Invoke-callable C API the way laszip3.dll does
    - that plain-C API is what let SCRLasZip be built as a P/Invoke shim in the first place, and PDAL
    simply has no equivalent to call into directly.

    pdal.exe gives no incremental progress for either use here - confirmed directly: neither pdal's own
    --progress <file> flag nor -v 4 (max useful verbosity) print anything between pipeline start and its
    own single "Wrote N points" line at the very end, because filters.sample/filters.delaunay both need
    the entire point view in memory before they can produce anything, so there's no per-point or
    per-chunk checkpoint to report mid-run. So progressCallback (if given) is just polled every ~1s to
    report elapsed time - not a fake percentage - so the UI visibly stays alive instead of showing a
    static message indistinguishable from a hang, and to notice a cancellation request: since
    progressCallback here is always ultimately SCR_CloudReviseDialog.set_run_progress, which raises
    PropellerRunCancelled itself when cancelled, this function's own job on that exception is only to
    kill the still-running pdal.exe process before letting it propagate, so a cancelled run doesn't leave
    an orphaned process behind.

    Raises RuntimeError (with pdal's own stderr) if pdal.exe exits non-zero."""
    proc = subprocess.Popen([pdalExePath, "pipeline", pipelinePath], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        startTime = timer()
        lastReport = startTime
        while proc.poll() is None:
            if progressCallback and (timer() - lastReport) > 1.0:
                lastReport = timer()
                elapsed = int(lastReport - startTime)
                progressCallback(statusText + "... " + str(elapsed) + "s elapsed", None)
            time.sleep(0.05)
    except:
        proc.terminate()
        proc.wait()
        raise

    if proc.returncode != 0:
        stderr = proc.stderr.read()
        raise RuntimeError("PDAL failed: " + (stderr.decode("utf-8", "ignore").strip() if stderr else "exit code " + str(proc.returncode)))


def pdal_resample_laz(pdalExePath, sourcePath, outputPath, spacing, progressCallback=None):
    """Thins a LAS/LAZ point cloud via QGIS's bundled PDAL (filters.sample - a minimum-radius, Poisson-
    disk-like filter, the closest native match to SCROctree.filter_min_spacing's own "octree" method
    here). Used for resampleMethod == "pdal" when triangulation is still done in-house (i.e. never - see
    _pdal_resample_and_triangulate_to_landxml, which is what "pdal" actually uses now - kept as a
    standalone building block since any future PDAL-thin-only path can still call it directly
    without pulling in triangulation).

    Raises RuntimeError (with pdal's own stderr) if pdal.exe exits non-zero. Writes outputPath on success."""
    pipeline = [
        {"type": "readers.las", "filename": sourcePath},
        {"type": "filters.sample", "radius": spacing},
        {"type": "writers.las", "filename": outputPath},
    ]
    pipelinePath = outputPath + ".pipeline.json"
    with open(pipelinePath, "w") as f:
        json.dump(pipeline, f)

    try:
        _run_pdal_pipeline(pdalExePath, pipelinePath, "Resampling with PDAL (" + os.path.basename(sourcePath) + ")", progressCallback)
    finally:
        try:
            os.remove(pipelinePath)
        except Exception:
            pass


def _read_ply_mesh(plyPath):
    """Parses an ASCII PLY (`format ascii 1.0`) written by PDAL's writers.ply into (points, faces).

    points: list of (x, y, z) - only the first three vertex properties are read; any extra per-vertex
    properties PDAL includes (e.g. offsettime) are ignored.
    faces: list of (i0, i1, i2), 0-based indices into points - a face line's leading vertex count is
    checked and only kept if it's exactly 3 (filters.delaunay only ever emits triangles for a 2.5D point
    cloud, so this is just a safety check, not an expected occurrence).

    No external PLY library used - the ASCII PLY subset PDAL writes here (single vertex + single face
    element, no binary, no list properties on vertices) is simple enough to hand-parse."""
    with open(plyPath, "r") as f:
        lines = f.read().split("\n")

    i = 0
    numVertices = numFaces = 0
    while lines[i].strip() != "end_header":
        line = lines[i].strip()
        if line.startswith("element vertex"):
            numVertices = int(line.split()[-1])
        elif line.startswith("element face"):
            numFaces = int(line.split()[-1])
        i += 1
    i += 1  # step past "end_header" itself

    points = []
    for _ in range(numVertices):
        parts = lines[i].split()
        i += 1
        points.append((float(parts[0]), float(parts[1]), float(parts[2])))

    faces = []
    for _ in range(numFaces):
        parts = lines[i].split()
        i += 1
        if int(parts[0]) == 3:
            faces.append((int(parts[1]), int(parts[2]), int(parts[3])))

    return points, faces


def _pdal_resample_and_triangulate_to_landxml(pdalExePath, lazPath, xmlPath, surfaceName, spacing, progressCallback=None):
    """PDAL-native alternative to the rest of laz_points_to_landxml_surface's own pipeline (SCRLasZip
    read -> Python resample -> Gem/AddVertex/ConstructSurface): runs ONE pdal.exe pipeline that both
    thins (filters.sample) and triangulates (filters.delaunay) the LAZ, then parses PDAL's own PLY mesh
    output straight into LandXML via _write_landxml_surface. Only used for resampleMethod == "pdal" -
    every other method keeps triangulating in-house via Gem, since routing through PDAL's own Delaunay
    (and paying for a PLY round-trip) is only worth it once PDAL is already doing the thinning too.

    PDAL has no LandXML writer of its own (checked: not in `pdal --drivers`'s writer list at all), so the
    mesh comes out as a plain ASCII PLY (vertex block + face block) via writers.ply, parsed here by
    _read_ply_mesh.

    writers.ply's own "precision" option defaults to just 3 - fine for small local coordinates, but
    Propeller LAZ files carry real-world UTM-style absolute coordinates (6-7 digit eastings/northings), so
    3 significant digits of headroom leaves only ~1m of resolution on easting and ~10m on northing (one
    more integer digit = one less digit of precision) once the ASCII writer switches to general/scientific
    formatting for a number that large. That's not a resampling artifact at all - it silently rounds
    surviving points onto an axis-dependent grid *after* thinning already picked them, which is exactly
    the "vertices land on an exact 1m/10m grid" pattern this was hunted down from (confirmed directly:
    the same pipeline against a real 46M-point Propeller LAZ produced coordinates like "6.9837e+06" at the
    default precision, and full "6983702.890100000426173" once bumped). 15 significant digits comfortably
    covers a double's own ~15-17 digit precision ceiling regardless of coordinate magnitude.

    PDAL's Delaunay has no equivalent to Gem.MaxOuterLength - it triangulates every point, gaps and all,
    so without a fix, long sliver triangles would reach across real gaps in the cloud all the way to the
    convex hull. This replicates the same fix the in-house path gets from Gem.MaxOuterLength: any face
    with an edge longer than 5x spacing is dropped (faces only - every point PDAL returns is still
    written, matching Gem's own behavior of keeping every vertex a surviving triangle still touches).

    Returns (pointCount, triangleCount), same contract as laz_points_to_landxml_surface itself."""
    plyPath = os.path.splitext(xmlPath)[0] + "_pdal_mesh.ply"
    pipeline = [
        {"type": "readers.las", "filename": lazPath},
        {"type": "filters.sample", "radius": spacing},
        {"type": "filters.delaunay"},
        {"type": "writers.ply", "filename": plyPath, "faces": True, "precision": 15},
    ]
    pipelinePath = plyPath + ".pipeline.json"
    with open(pipelinePath, "w") as f:
        json.dump(pipeline, f)

    try:
        _run_pdal_pipeline(pdalExePath, pipelinePath, "Resampling and triangulating with PDAL (" + os.path.basename(lazPath) + ")", progressCallback)

        if progressCallback:
            progressCallback("Reading PDAL mesh...", 70)
        points, rawFaces = _read_ply_mesh(plyPath)
    finally:
        for p in (pipelinePath, plyPath):
            try:
                os.remove(p)
            except Exception:
                pass

    maxEdge2 = (spacing * 5) ** 2
    faces = []
    for i0, i1, i2 in rawFaces:
        p0, p1, p2 = points[i0], points[i1], points[i2]
        d01 = (p0[0]-p1[0])**2 + (p0[1]-p1[1])**2 + (p0[2]-p1[2])**2
        d12 = (p1[0]-p2[0])**2 + (p1[1]-p2[1])**2 + (p1[2]-p2[2])**2
        d20 = (p2[0]-p0[0])**2 + (p2[1]-p0[1])**2 + (p2[2]-p0[2])**2
        if d01 <= maxEdge2 and d12 <= maxEdge2 and d20 <= maxEdge2:
            faces.append((i0, i1, i2))

    if progressCallback:
        progressCallback("Writing LandXML...", 90)
    _write_landxml_surface(xmlPath, surfaceName, "pdal", spacing, points, faces)

    if progressCallback:
        progressCallback(None, 100)

    return len(points), len(faces)


def _write_landxml_surface(xmlPath, surfaceName, resampleMethod, resampleSpacing, points, faces):
    """Hand-writes a minimal LandXML 1.2 [Surface] (Pnts + Faces, surfType=TIN) - shared by both
    triangulation paths in laz_points_to_landxml_surface (Gem-based in-house methods, and PDAL's own
    Delaunay for resampleMethod == "pdal"), so the file format itself only needs writing once. Bypasses
    TBC's own LandXmlExporter, which needs a real project entity selection (ICollectionOfEntities) and so
    can't run against in-memory-only point/face data.

    points: dense list of (x, y, z) real-world coordinates - point i becomes LandXML point id i+1.
    faces: list of (i0, i1, i2), 0-based indices into points."""
    folder = os.path.dirname(xmlPath)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    if resampleSpacing:
        methodLabel = {"voxel": "voxel grid", "octree": "octree min-spacing", "grid": "regular grid", "pdal": "PDAL"}.get(resampleMethod, resampleMethod)
        fullSurfaceName = surfaceName + " (" + methodLabel + ", " + ("%g" % resampleSpacing) + "m spacing)"
    else:
        fullSurfaceName = surfaceName

    with open(xmlPath, "w") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write('<LandXML xmlns="http://www.landxml.org/schema/LandXML-1.2" version="1.2" date="' + datetime.now().strftime("%Y-%m-%d") + '" time="' + datetime.now().strftime("%H:%M:%S") + '">\n')
        f.write('  <Units>\n')
        f.write('    <Metric areaUnit="squareMeter" linearUnit="meter" volumeUnit="cubicMeter" temperatureUnit="celsius" pressureUnit="mmHG"/>\n')
        f.write('  </Units>\n')
        f.write('  <Surfaces>\n')
        f.write('    <Surface name="' + fullSurfaceName + '">\n')
        f.write('      <Definition surfType="TIN">\n')
        f.write('        <Pnts>\n')
        for i, (x, y, z) in enumerate(points):
            # LandXML point order is northing(Y) easting(X) elevation(Z), not X Y Z - a common gotcha.
            # Rounded to millimeters rather than repr()'d - repr() can emit up to ~17 significant digits
            # of floating-point noise, which is both meaningless precision for survey data (doubly so
            # here, since the in-house path's underlying coordinates already only carry float32-level
            # precision - see SCRLasZip.read_points' relative=True mode) and unnecessary bytes for TBC's
            # importer to parse across hundreds of thousands of points
            f.write('          <P id="' + str(i + 1) + '">' + ("%.3f" % y) + ' ' + ("%.3f" % x) + ' ' + ("%.3f" % z) + '</P>\n')
        f.write('        </Pnts>\n')
        f.write('        <Faces>\n')
        for i0, i1, i2 in faces:
            f.write('          <F>' + str(i0 + 1) + ' ' + str(i1 + 1) + ' ' + str(i2 + 1) + '</F>\n')
        f.write('        </Faces>\n')
        f.write('      </Definition>\n')
        f.write('    </Surface>\n')
        f.write('  </Surfaces>\n')
        f.write('</LandXML>\n')


def laz_points_to_landxml_surface(lazPath, xmlPath, surfaceName="Surface", progressCallback=None, resampleSpacing=None, resampleMethod="voxel", expectedPointCount=None, pdalExePath=None):
    """Reads a LAS/LAZ file via SCRLasZip (no PointCloudDatabase, no project import), triangulates
    the points in-memory with a bare Gem(Gem.ModelType.Dtm) (same technique SCR_VPath uses for its
    concave hull - see SCR_VPath.py's _alpha_hull - a Gem triangulates without ever being attached to
    a worldview surface), and hand-writes the result as a minimal LandXML 1.2 [Surface] - bypassing
    TBC's own LandXmlExporter, which needs a real project entity selection (ICollectionOfEntities)
    and so can't run against an in-memory-only Gem.

    resampleSpacing, if given, thins the points before triangulating via resampleMethod: "voxel"
    (resample_points_voxel_grid - fast, density-target, follows the cloud's own point distribution),
    "octree" (SCROctree.filter_min_spacing - slower, true minimum spacing, also follows the cloud's own
    distribution), "grid" (resample_points_grid_2d - regular XY grid, nearest real point per cell by
    2D distance; usually the fastest of the three for a dense, roughly uniform cloud, and gives
    perfectly uniform output spacing rather than following the input's density), or "pdal" (hands the
    whole job - thinning AND triangulating - to QGIS's bundled pdal.exe via
    _pdal_resample_and_triangulate_to_landxml, skipping SCRLasZip/Gem entirely; typically far faster than
    all three in-house methods at real drone-LAZ point counts, but requires pdalExePath to point at a
    real pdal.exe). None (the default) skips resampling entirely.

    pdalExePath is only used when resampleMethod == "pdal". If it's missing or doesn't exist, this
    silently falls back to "voxel" instead of failing the whole run - a schedule entry saved with "pdal"
    in one session can easily outlive that session's QGIS-PDAL configuration (folder unset, QGIS
    uninstalled, etc.), and a stale setting like that shouldn't turn into a hard failure days later.

    expectedPointCount, if given, is passed straight through to SCRLasZip.read_points as a hint to
    pre-size its output arrays - see that method's docstring. Only ever a speed hint, never trusted over
    what the file actually contains. Unused for resampleMethod == "pdal", which never calls SCRLasZip.

    Returns (pointCount, triangleCount).
    """
    def fmt(n):
        # IronPython ints are real .NET integers - ToString("N0") gives thousands separators
        # (str.format's "{:,}" doesn't work the same way in IronPython, hence going straight to .NET)
        return n.ToString("N0")

    def report_read_progress(i, n):
        progressCallback("Reading points... " + fmt(i) + "/" + fmt(n), int(i * 30 / n))
        return False  # never cancel

    if resampleMethod == "pdal" and not (pdalExePath and os.path.isfile(pdalExePath)):
        if progressCallback:
            progressCallback("PDAL not available - falling back to voxel resampling...", 0)
        resampleMethod = "voxel"

    if resampleMethod == "pdal" and resampleSpacing:
        return _pdal_resample_and_triangulate_to_landxml(pdalExePath, lazPath, xmlPath, surfaceName, resampleSpacing, progressCallback)

    if progressCallback:
        progressCallback("Reading " + os.path.basename(lazPath) + "...", 0)
    # relative=True reads into flat array.array('f') coordinate arrays offset from the first point,
    # instead of a Python list of (x, y, z) tuples - at tens of millions of points, avoiding one boxed
    # float object per coordinate plus one tuple object per point is what actually matters for memory,
    # not just the smaller float width (see SCRLasZip.read_points' docstring). Kept in this offset-
    # relative form all the way through triangulation below - distance-based resampling and triangle
    # connectivity are both translation-invariant, so nothing needs real coordinates until the LandXML
    # <Pnts> are actually written.
    offsetX, offsetY, offsetZ, xs, ys, zs, relativeBounds = SCRLasZip.read_points(lazPath, progressCallback=report_read_progress if progressCallback else None, relative=True, expectedCount=expectedPointCount)
    pointCount = len(xs)

    relativePoints = None
    if resampleSpacing:
        def report_resample_progress(i, n):
            progressCallback("Resampling... " + fmt(i) + "/" + fmt(n), 30 + int(i * 10 / n))
            return False  # never cancel

        # zip(xs, ys, zs) would do this same job in one call, but it's an atomic builtin with no hook for
        # progress in between - for tens of millions of points that's still a real multi-second pause,
        # just with a label sitting frozen in front of it instead of no label at all. Building the list by
        # hand is the same total work, just with a chance to report progress along the way.
        if progressCallback:
            progressCallback("Preparing " + fmt(pointCount) + " points for resampling...", 30)
            relativePoints = []
            lastReport = timer()
            for i in range(pointCount):
                relativePoints.append((xs[i], ys[i], zs[i]))
                if (timer() - lastReport) > 1.0:
                    progressCallback("Preparing " + fmt(i) + "/" + fmt(pointCount) + " points for resampling...", 30 + int(i * 2 / pointCount))
                    lastReport = timer()
        else:
            relativePoints = zip(xs, ys, zs)  # resample_points_voxel_grid/SCROctree need real (x, y, z) tuples to work with

        if resampleMethod == "octree":
            # bounds were already tracked for free while reading above (see SCRLasZip.read_points'
            # bounds return value) - handing them to filter_min_spacing skips its own separate full pass
            # over the points computing the exact same thing, which used to be slow enough at tens of
            # millions of points to need its own "Preparing octree..." progress message
            if progressCallback:
                progressCallback("Resampling (octree, min spacing) " + fmt(pointCount) + " points...", 32)
            relativePoints = SCROctree().filter_min_spacing(relativePoints, resampleSpacing, progressCallback=report_resample_progress if progressCallback else None, bounds=relativeBounds)
        elif resampleMethod == "grid":
            # regular-grid resample - fast for a dense, roughly uniform cloud (typical for LAZ) since
            # grid cell count is usually far smaller than point count at a coarse target spacing; see
            # resample_points_grid_2d's docstring for the full reasoning
            if progressCallback:
                progressCallback("Resampling (regular grid) " + fmt(pointCount) + " points...", 32)
            relativePoints = resample_points_grid_2d(relativePoints, resampleSpacing, bounds=relativeBounds, progressCallback=report_resample_progress if progressCallback else None)
        else:
            if progressCallback:
                progressCallback("Resampling " + fmt(pointCount) + " points...", 32)
            relativePoints = resample_points_voxel_grid(relativePoints, resampleSpacing, progressCallback=report_resample_progress if progressCallback else None)
        pointCount = len(relativePoints)

    if progressCallback:
        progressCallback("Triangulating " + fmt(pointCount) + " points...", 40)

    gem = Gem(Gem.ModelType.Dtm)
    if resampleSpacing:
        # trims outer triangle edges longer than this during ConstructSurface() (see SCR_VPath.py's
        # _alpha_hull for the same technique) - without it, sparse/irregular areas of the cloud get
        # stitched together by long sliver triangles reaching all the way to the convex hull instead of
        # leaving a hole. 5x the resample spacing is generous enough that legitimately adjacent points at
        # the target density still connect, while edges much longer than that (real gaps) don't.
        gem.MaxOuterLength = resampleSpacing * 5
    lastReport = timer()
    if relativePoints is not None:
        for i, (x, y, z) in enumerate(relativePoints):
            gem.AddVertex(False, GemVertexType.Control, 0, Point3D(x + offsetX, y + offsetY, z + offsetZ))
            if progressCallback and (timer() - lastReport) > 1.0:
                progressCallback("Adding vertices... " + fmt(i) + "/" + fmt(pointCount), 40 + int(i * 30 / pointCount))
                lastReport = timer()
    else:
        for i in range(pointCount):
            gem.AddVertex(False, GemVertexType.Control, 0, Point3D(xs[i] + offsetX, ys[i] + offsetY, zs[i] + offsetZ))
            if progressCallback and (timer() - lastReport) > 1.0:
                progressCallback("Adding vertices... " + fmt(i) + "/" + fmt(pointCount), 40 + int(i * 30 / pointCount))
                lastReport = timer()

    gem.ConstructSurface()

    if progressCallback:
        progressCallback("Writing LandXML...", 80)

    # precompute into flat Python lists before the write loop - repeated marshaled Gem calls inside a
    # hot loop are the dominant cost here, not the file I/O (same lesson as the GDAL raster reads above).
    # Also remaps old (0-based, sparse - a Gem can have gaps) vertex indices to new (0-based, dense)
    # indices into densePoints as it goes, since _write_landxml_surface expects a dense points list.
    landxmlIndex = {}
    densePoints = []
    for i in range(gem.NumberOfVertices):
        if not gem.IsVertexPresent(i):
            continue
        landxmlIndex[i] = len(densePoints)
        p = gem.GetVertexPoint(i)
        densePoints.append((p.X, p.Y, p.Z))

    faces = []
    for t in range(gem.NumberOfTriangles):
        if not gem.IsTrianglePresent(t):
            continue
        v0, v1, v2 = gem.GetTriangleVertex(t, 0), gem.GetTriangleVertex(t, 1), gem.GetTriangleVertex(t, 2)
        faces.append((landxmlIndex[v0], landxmlIndex[v1], landxmlIndex[v2]))

    _write_landxml_surface(xmlPath, surfaceName, resampleMethod, resampleSpacing, densePoints, faces)

    if progressCallback:
        progressCallback(None, 100)

    return len(densePoints), len(faces)


def join_folder_and_file(folderPath, fileName):
    if folderPath.endswith("/"):
        return folderPath + fileName
    return folderPath + "/" + fileName


def collect_civillo_files(node, results):
    """Recursively walks a /layer-directory response tree, collecting {"display", "layerId", "folderPath",
    "layerName"} dicts for every file. folderPath/layerName are kept separate (alongside the combined
    "display" string) so callers can render/store them differently, e.g. coloring the folder path and
    the file name differently in a list item."""
    folderPath = node.get("path", "/")
    for layer in node.get("layers", []):
        results.append({
            "display": join_folder_and_file(folderPath, layer["layerName"]),
            "layerId": layer["layerId"],
            "folderPath": folderPath,
            "layerName": layer["layerName"],
        })
    for d in node.get("directories", []):
        collect_civillo_files(d, results)


def dump_civillo_layer_directory(tree, configFolder):
    """Writes the raw, unmodified /layer-directory JSON response next to credentials.json - a temporary
    diagnostic aid for checking exactly what Civillo's API returns for a project (e.g. whether a terrain/
    DTM/surface layer shows up in the "layers" array at all, since collect_civillo_files above doesn't
    filter by any type and Civillo's own documented schema has no type field to filter on in the first
    place - if a terrain layer is missing from our own file list, this dump settles whether that's because
    the API never returned it, rather than something being dropped on our side). Overwrites on every
    fetch; safe to ignore/delete, not read by anything else in this macro."""
    try:
        dumpPath = os.path.join(configFolder, "civillo_layerdirectory_dump.json")
        with open(dumpPath, "w") as f:
            f.write(json.dumps(tree, indent=2, sort_keys=True))
    except Exception:
        pass  # purely diagnostic - never let a failure here interrupt the real file-list load


def get_civillo_unassigned_layers(civilloClient, orgNickname, projectId, directoryFiles):
    """Returns every layer from the flat GET /layers endpoint whose layerId never showed up while
    walking /layer-directory (directoryFiles, already collected via collect_civillo_files) - by
    elimination, these are layers that exist in the project but aren't organized into any folder, which
    is exactly the situation a terrain/DTM/surface layer is always in: confirmed by hand (see
    conversation) that terrain layers never appear in /layer-directory at all, no matter what, while
    they do show up in the plain /layers listing - and neither that listing nor the single-layer detail
    endpoint (/layers/{layerId}) expose any type/category field to filter on directly, so this
    elimination is the only way to find them via the API.

    Not a pure "terrain-only" filter - a CAD layer can also be deliberately left unassigned to any
    directory folder (kept "in the background") without being a terrain upload at all, and a couple of
    other special cases (e.g. a WMTS basemap overlay layer, which has no uploaded file at all) end up
    in this same elimination set too. So this is presented as "Unassigned Layers" generally, not
    labelled "Terrain" - technically accurate either way, and avoids mislabeling one of those other cases.

    Returns a list of the raw layer dicts from GET /layers (id/layerId/name/fromFileName/etc.) for every
    layerId not present in directoryFiles. Raises on an API failure - same as the /layer-directory call
    this supplements, left for the caller to catch alongside it."""
    directoryIds = set(f["layerId"] for f in directoryFiles)
    allLayers = civilloClient.get("/" + orgNickname + "/projects/" + str(projectId) + "/layers")
    return [l for l in allLayers if l.get("layerId") not in directoryIds]


def Setup(cmdData, macroFileFolder):
    cmdData.Key = "SCR_CloudRevise"
    cmdData.CommandName = "SCR_CloudRevise"
    cmdData.Caption = "_SCR_CloudRevise"
    #cmdData.UIForm = "SCR_CloudRevise"      # left disabled - this is a fully independent floating window, not a TBC-managed dialog
                                                        # if you enable or disable this line, you MUST restart TBC
    cmdData.HelpFile = "Macros.chm"
    cmdData.HelpTopic = "0"

    try:
        cmdData.DefaultTabKey = "SCR ImExport/DTM/Subgrade"
        cmdData.DefaultTabGroupKey = "Update/Transfer"
        cmdData.ShortCaption = "Cloud Revise"
        cmdData.DefaultRibbonToolSize = 3 # Default=0, ImageOnly=1, Normal=2, Large=3
        cmdData.EnableNoProject       = True
        
        cmdData.Version = 1.2
        cmdData.MacroAuthor = "SCR"
        cmdData.MacroInfo = r""

        cmdData.ToolTipTitle = "CloudRevise"
        cmdData.ToolTipTextFormatted = "transfer new files and revise Civillo layers"

    except:
        pass
    try:
        b = Bitmap (macroFileFolder + "\\" + cmdData.Key + ".png") # we have to include a icon revision, otherwise TBC might not show the new one
        cmdData.ImageSmall = b
    except:
        pass

def Execute(cmd, currentProject, macroFileFolder, parameters):
    form = SCR_CloudReviseDialog(currentProject, macroFileFolder).Show()
    return
    # .Show() - is non modal - you can interact with the drawing window
    # .ShowDialog() - is modal - you CAN NOT interact with the drawing window


class SCR_CloudReviseDialog(Window): # this inherits from the WPF Window control - a fully independent floating window
    def __init__(self, currentProject, macroFileFolder):

        with StreamReader(macroFileFolder + r"\SCR_CloudRevise.xaml") as s:
            wpf.LoadComponent(self, s)

        ElementHost.EnableModelessKeyboardInterop(self)

        self.currentProject = currentProject
        self.macroFileFolder = macroFileFolder

        self.orgCombo.SelectionChanged += self.org_selection_changed
        self.projectCombo.SelectionChanged += self.project_selection_changed
        self.reloadOrgsBtn.Click += self.reload_orgs_clicked
        self.addScheduleBtn.Click += self.add_schedule_clicked
        self.runSchedulesBtn.Click += self.run_schedules_clicked
        self.helpBtn.Click += self.help_clicked
        self.configPathLabel.MouseLeftButtonDown += self.config_path_clicked
        self.openSettingsFolderLabel.MouseLeftButtonDown += self.open_settings_folder_clicked

        self.qgisFolderBox.TextChanged += self.qgis_folder_changed
        self.browseQgisFolderBtn.Click += self.browse_qgis_folder_clicked

        self.propReloadBtn.Click += self.reload_propeller_orgs_clicked
        self.propOrgCombo.SelectionChanged += self.propeller_org_selection_changed
        self.propSiteFilterBox.TextChanged += self.propeller_site_filter_changed
        self.propSiteCombo.SelectionChanged += self.propeller_site_selection_changed
        self.propSurveyCombo.SelectionChanged += self.propeller_survey_selection_changed

        self.addPropellerScheduleBtn.Click += self.add_propeller_schedule_clicked
        self.runPropellerSchedulesBtn.Click += self.run_propeller_schedules_clicked
        self.cancelPropellerRunBtn.Click += self.cancel_propeller_run_clicked

        self._propellerCancelRequested = False
        self._closeRequestedDuringRun = False
        self.propellerSites = []

        # debounces the site combo's filter-driven reselection (see propeller_site_filter_changed) - a
        # fast typist composing a longer search term would otherwise trigger a live get_surveys() call
        # for every intermediate site the filter briefly lands on along the way
        from System.Windows.Threading import DispatcherTimer
        self._propSiteFilterDebounceTimer = DispatcherTimer()
        self._propSiteFilterDebounceTimer.Interval = System.TimeSpan.FromMilliseconds(250)
        self._propSiteFilterDebounceTimer.Tick += self.prop_site_filter_debounce_elapsed
        self.propellerSelectedFiles = []
        self.propellerSurveys = []  # full survey dicts for the currently selected site, date_captured included - see propeller_site_selection_changed
        self._lastProcessedPropSiteId = None

        self.qgisFolderBox.Text = OptionsManager.GetString("SCR_CloudRevise.qgisfolder", "")
        self._update_pdal_status()

        self.Loaded += self.SetDefaultOptions
        self.Closing += self.SaveOptions

        configPath = self.get_config_path()
        self.configPathLabel.Text = "config: " + configPath
        credentials = self.load_or_create_config(configPath)
        self.civilloConfig = credentials["civillo"]
        self.propellerConfig = credentials["propeller"]
        self.civilloClient = CivilloClient(self.civilloConfig)
        self.propellerClient = PropellerClient(self.propellerConfig)
        self.trimbleConnectClient = TrimbleConnectClient()

        self.propellerDownloadCache = PropellerDownloadCache(self.get_propeller_cache_folder())
        self.propellerDownloadCache.cleanup()

        self.schedules = self.load_schedules()
        self.backfill_schedule_names()
        self.refresh_schedules_ui()

        self.propellerSchedules = self.load_propeller_schedules()
        self.refresh_propeller_schedules_ui()
        self.reload_propeller_orgs_clicked(None, None)

        if self.is_placeholder_config(self.civilloConfig):
            self.error.Content = "Edit the config file above with your Civillo API key/secret, then click 'Reload Organizations'."
        else:
            self.load_organizations()


    # ---------- QGIS-PDAL cloud resampling ----------

    def _update_pdal_status(self):
        folder = self.qgisFolderBox.Text.strip()
        self.pdalExePath = os.path.join(folder, "bin", "pdal.exe") if folder else None
        self.pdalAvailable = bool(self.pdalExePath) and os.path.isfile(self.pdalExePath)
        if not folder:
            self.pdalStatusLabel.Text = ""
            self.pdalStatusLabel.Foreground = SolidColorBrush(Colors.Gray)
        elif self.pdalAvailable:
            self.pdalStatusLabel.Text = "Found: " + self.pdalExePath
            self.pdalStatusLabel.Foreground = SolidColorBrush(Colors.DarkGreen)
        else:
            # a QGIS update can silently move pdal.exe to a new version-numbered folder (e.g.
            # "QGIS 3.36.2" -> "QGIS 3.40.0") - flagging this in red is what actually gets it noticed,
            # rather than every subsequent LAZ run just quietly falling back to voxel resampling
            self.pdalStatusLabel.Text = "pdal.exe not found under " + os.path.join(folder, "bin") + " - check the QGIS folder (a QGIS update may have changed the install path)"
            self.pdalStatusLabel.Foreground = SolidColorBrush(Colors.Red)

    def qgis_folder_changed(self, sender, e):
        OptionsManager.SetValue("SCR_CloudRevise.qgisfolder", self.qgisFolderBox.Text)
        self._update_pdal_status()

    def browse_qgis_folder_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        dlg.Description = "Select the main QGIS install folder (contains a 'bin' subfolder with pdal.exe)"
        if self.qgisFolderBox.Text.strip():
            dlg.SelectedPath = self.qgisFolderBox.Text.strip()
        if dlg.ShowDialog() == DialogResult.OK:
            self.qgisFolderBox.Text = dlg.SelectedPath  # fires qgis_folder_changed, which persists + re-validates
            OptionsManager.SaveOptions()  # a folder picked through Browse is a discrete, infrequent action - flush now


    # ---------- window position/size ----------

    def SetDefaultOptions(self, sender, e):
        SCROptions.LoadWindowState(self, "SCR_CloudRevise", default_width=300, default_height=460)

    def SaveOptions(self, sender, e):
        SCROptions.SaveWindowState(self, "SCR_CloudRevise")

        # a Propeller run (LAZ resampling included) is still executing synchronously on this same UI
        # thread - closing the window doesn't interrupt it by itself, since nothing was checking for
        # that, only the Cancel button ever set the flag the run actually watches. Request cancellation
        # the same way Cancel does, but defer the actual close until the run has had a chance to unwind
        # cleanly (_end_cancellable_run re-issues Close() once it has) - closing right now instead would
        # dispose this window while the still-running code underneath tries to keep touching it.
        if self.cancelPropellerRunBtn.IsEnabled:
            self._propellerCancelRequested = True
            self._closeRequestedDuringRun = True
            e.Cancel = True
            return

        self._propSiteFilterDebounceTimer.Stop()  # a pending Tick would otherwise still fire against this now-closed window
        OptionsManager.SaveOptions()  # flush window size/state to disk now rather than relying on TBC's own exit-time flush


    # ---------- org/project selection persistence ----------

    def get_saved_selection(self):
        nickname = OptionsManager.GetString("SCR_CloudRevise.selectedorgnickname", "")
        projectId = OptionsManager.GetString("SCR_CloudRevise.selectedprojectid", "")
        return nickname, projectId

    def save_selection(self, nickname, projectId):
        OptionsManager.SetValue("SCR_CloudRevise.selectedorgnickname", str(nickname) if nickname is not None else "")
        OptionsManager.SetValue("SCR_CloudRevise.selectedprojectid", str(projectId) if projectId is not None else "")


    # ---------- config handling ----------

    def get_config_path(self):
        appdata = os.environ.get("APPDATA")
        folder = os.path.join(appdata, "SCR Macros", "SCR_CloudRevise")
        return os.path.join(folder, "credentials.json")

    def get_legacy_config_path(self):
        return os.path.join(os.path.dirname(self.get_config_path()), "civillo_config.json")

    def load_or_create_config(self, path):
        folder = os.path.dirname(path)
        if not os.path.exists(folder):
            os.makedirs(folder)

        if not os.path.exists(path):
            legacyPath = self.get_legacy_config_path()
            if os.path.exists(legacyPath):
                # upgrading from the old civillo-only config file - carry the key/secret over into
                # the new multi-system shape, then remove the old file so it isn't mistaken for current
                with open(legacyPath, "r") as f:
                    legacy = json.load(f)
                credentials = {
                    "civillo": {
                        "api_key": legacy.get("api_key", CIVILLO_PLACEHOLDER_KEY),
                        "api_secret": legacy.get("api_secret", CIVILLO_PLACEHOLDER_SECRET),
                    },
                    "propeller": {
                        "access_token": PROPELLER_PLACEHOLDER_TOKEN,
                        "_note": "access_token is the raw token only - do NOT include a 'Bearer ' prefix, that is added automatically",
                    },
                }
                with open(path, "w") as f:
                    f.write(json.dumps(credentials, indent=2))
                os.remove(legacyPath)
                return credentials

            credentials = {
                "civillo": {
                    "api_key": CIVILLO_PLACEHOLDER_KEY,
                    "api_secret": CIVILLO_PLACEHOLDER_SECRET,
                },
                "propeller": {
                    "access_token": PROPELLER_PLACEHOLDER_TOKEN,
                },
            }
            with open(path, "w") as f:
                f.write(json.dumps(credentials, indent=2))
            return credentials

        with open(path, "r") as f:
            return json.load(f)

    def is_placeholder_config(self, cfg):
        return cfg.get("api_key", "") == CIVILLO_PLACEHOLDER_KEY or cfg.get("api_secret", "") == CIVILLO_PLACEHOLDER_SECRET

    def config_path_clicked(self, sender, e):
        subprocess.Popen(["notepad.exe", self.get_config_path()])

    def open_settings_folder_clicked(self, sender, e):
        folder = os.path.dirname(self.get_config_path())
        if not os.path.isdir(folder):
            os.makedirs(folder)
        subprocess.Popen(["explorer.exe", folder])


    # ---------- Propeller download cache ----------

    def get_propeller_cache_folder(self):
        return os.path.join(os.path.dirname(self.get_config_path()), "Propeller-Cache")


    # ---------- schedules persistence ----------

    def get_schedules_path(self):
        return os.path.join(os.path.dirname(self.get_config_path()), "schedules.json")

    def load_schedules(self):
        path = self.get_schedules_path()
        if not os.path.exists(path):
            return []
        with open(path, "r") as f:
            return json.load(f)

    def save_schedules(self):
        path = self.get_schedules_path()
        folder = os.path.dirname(path)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, "w") as f:
            f.write(json.dumps(self.schedules, indent=2))

    def backfill_schedule_names(self):
        # (Re)resolve orgName/projectName for every schedule entry from the live API, so the Schedules
        # tab group header always matches the plain names shown in the Settings tab dropdowns - this
        # both fills in older entries that only have the raw orgNickname/projectId, and corrects any
        # entries that briefly picked up an "extended" name format from a since-reverted change.
        nicknames = set(entry["orgNickname"] for entry in self.schedules if entry.get("orgNickname"))
        if not nicknames:
            return

        orgNameByNickname = {}
        try:
            orgs = self.civilloClient.get("/applications")
            for org in orgs:
                orgNameByNickname[org["nickname"]] = org["name"]
        except Exception:
            return  # API not reachable right now (e.g. placeholder config) - leave as-is for next time

        projectNameByKey = {}
        for nickname in nicknames:
            try:
                projects = self.civilloClient.get("/" + nickname + "/projects")
                for proj in projects:
                    projectNameByKey[(nickname, proj["id"])] = proj["name"]
            except Exception:
                pass

        changed = False
        for entry in self.schedules:
            nickname = entry.get("orgNickname")
            if nickname in orgNameByNickname and entry.get("orgName") != orgNameByNickname[nickname]:
                entry["orgName"] = orgNameByNickname[nickname]
                changed = True
            key = (nickname, entry.get("projectId"))
            if key in projectNameByKey and entry.get("projectName") != projectNameByKey[key]:
                entry["projectName"] = projectNameByKey[key]
                changed = True

        if changed:
            self.save_schedules()


    # ---------- Propeller -> Civillo/Trimble Connect schedules persistence ----------
    # kept as a separate list/file from the Local -> Civillo schedules above, since a Propeller sync
    # entry's shape (source survey + Civillo-or-Trimble-Connect target) will differ once that flow is built

    def get_propeller_schedules_path(self):
        return os.path.join(os.path.dirname(self.get_config_path()), "propeller_schedules.json")

    def load_propeller_schedules(self):
        path = self.get_propeller_schedules_path()
        if not os.path.exists(path):
            return []
        with open(path, "r") as f:
            return json.load(f)

    def save_propeller_schedules(self):
        path = self.get_propeller_schedules_path()
        folder = os.path.dirname(path)
        if not os.path.exists(folder):
            os.makedirs(folder)
        # sort_keys groups every entry's fields by their shared "civillo"/"propeller"/"trimbleConnect"
        # prefix for free (ok_clicked builds each entry's dict with keys interleaved roughly by platform
        # already, but "enabled" gets added after the fact by add/edit_propeller_schedule_clicked, and
        # civilloEnabled/trimbleConnectEnabled/label get appended at the very end rather than grouped with
        # their own section - alphabetical order is a simple, always-consistent fix without having to
        # maintain a hand-written key order that'd need updating every time a field gets added)
        with open(path, "w") as f:
            f.write(json.dumps(self.propellerSchedules, indent=2, sort_keys=True))

    def refresh_propeller_schedules_ui(self):
        self.propellerSchedulesPanel.Children.Clear()

        # a schedule entry belongs to one site (propellerSiteId) but no fixed survey any more, so this
        # list only ever shows entries for whatever site is currently selected - switching sites switches
        # which entries are visible/runnable, same as switching surveys already changes which files
        # run_propeller_schedules_clicked can find a name match for
        siteItem = self.propSiteCombo.SelectedItem
        currentSiteId = siteItem.Tag if siteItem is not None else None
        indices = [i for i in range(len(self.propellerSchedules)) if self.propellerSchedules[i].get("propellerSiteId") == currentSiteId]
        indices.sort(key=lambda i: self.propeller_schedule_group_key(self.propellerSchedules[i]))

        lastGroupKey = None
        for i in indices:
            entry = self.propellerSchedules[i]
            groupKey = self.propeller_schedule_group_key(entry)
            if groupKey != lastGroupKey:
                self.propellerSchedulesPanel.Children.Add(self.build_propeller_schedule_group_header(entry))
                lastGroupKey = groupKey
            self.propellerSchedulesPanel.Children.Add(self.build_propeller_schedule_row(i))

    def propeller_schedule_group_key(self, entry):
        # groups by whichever target(s) the entry has - Civillo org/project, Trimble Connect region/project,
        # or both, since an entry can now target either or both
        return (
            entry.get("civilloOrgName") or entry.get("civilloOrgNickname") or "",
            entry.get("civilloProjectName") or str(entry.get("civilloProjectId") or ""),
            entry.get("trimbleConnectRegionName") or "",
            entry.get("trimbleConnectProjectName") or "",
        )

    def propeller_conversion_description(self, entry, platformPrefix):
        """Short one-line summary of what happens to this entry's source file for one platform target,
        e.g. "TIFF -> 0.05 m/px JPEG" (orthophoto resample) or "LAZ -> 1.3 m LandXML" (LAZ resample +
        triangulate). platformPrefix is "civillo" or "trimbleConnect" - matches the same prefix used for
        that platform's own per-entry keys (civilloPixelSizeMeters/trimbleConnectPixelSizeMeters, etc.),
        since the two platforms can be configured differently for the same entry. Returns None if there's
        nothing meaningful to summarize (a raw passthrough - no pixel size and no LAZ spacing set)."""
        sourceFormat = (entry.get("propellerFileFormat") or "").upper()
        pixelSize = entry.get(platformPrefix + "PixelSizeMeters")
        if pixelSize is not None:
            outputFormat = entry.get(platformPrefix + "OutputFormat") or "TIFF"
            return sourceFormat + " -> " + ("%g" % pixelSize) + " m/px " + outputFormat
        lazSpacing = entry.get(platformPrefix + "LazResampleSpacing")
        if lazSpacing is not None:
            return sourceFormat + " -> " + ("%g" % lazSpacing) + " m LandXML"
        return None

    def build_propeller_schedule_group_header(self, entry):
        # the description reflects this one entry (whichever started the group - see
        # refresh_propeller_schedules_ui) - a group can hold other entries targeting the same org/project
        # with different settings, so this is a representative hint for the group, not a guarantee every
        # entry in it matches
        labels = []
        if entry.get("civilloOrgName") or entry.get("civilloOrgNickname") or entry.get("civilloProjectName") or entry.get("civilloProjectId"):
            civilloDesc = self.propeller_conversion_description(entry, "civillo")
            labels.append("Civillo:" + (" " + civilloDesc if civilloDesc else ""))
        if entry.get("trimbleConnectRegionName") or entry.get("trimbleConnectProjectName"):
            tcDesc = self.propeller_conversion_description(entry, "trimbleConnect")
            labels.append("Connect:" + (" " + tcDesc if tcDesc else ""))
        if not labels:
            labels.append("Unknown Target")

        header = TextBlock()
        header.Text = " | ".join(labels)
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 8, 0, 4)
        return header

    def build_propeller_schedule_row(self, index):
        entry = self.propellerSchedules[index]

        row = Grid()
        row.Margin = Thickness(0, 0, 0, 8)

        colCheck = ColumnDefinition()
        colCheck.Width = GridLength(1, GridUnitType.Auto)

        colText = ColumnDefinition()
        colText.Width = GridLength(1, GridUnitType.Star)

        colEdit = ColumnDefinition()
        colEdit.Width = GridLength(1, GridUnitType.Auto)

        colDelete = ColumnDefinition()
        colDelete.Width = GridLength(1, GridUnitType.Auto)

        row.ColumnDefinitions.Add(colCheck)
        row.ColumnDefinitions.Add(colText)
        row.ColumnDefinitions.Add(colEdit)
        row.ColumnDefinitions.Add(colDelete)

        chk = CheckBox()
        chk.IsChecked = bool(entry.get("enabled", True))
        Grid.SetColumn(chk, 0)
        chk.Checked += lambda s, e, i=index: self.propeller_schedule_enabled_toggled(i, True)
        chk.Unchecked += lambda s, e, i=index: self.propeller_schedule_enabled_toggled(i, False)
        row.Children.Add(chk)

        propellerRun = Run(entry.get("propellerFileName", ""))
        propellerRun.Foreground = SolidColorBrush(Colors.DarkGreen)

        text = TextBlock()
        text.TextWrapping = TextWrapping.Wrap
        text.Margin = Thickness(6, 0, 6, 0)

        propellerHeader = Run("Propeller")
        propellerHeader.FontWeight = FontWeights.Bold
        propellerHeader.FontStyle = FontStyles.Italic
        text.Inlines.Add(propellerHeader)
        text.Inlines.Add(LineBreak())
        text.Inlines.Add(propellerRun)

        if entry.get("civilloDisplay") or entry.get("civilloCreateNewLayer"):
            civilloHeader = Run("Civillo")
            civilloHeader.FontWeight = FontWeights.Bold
            civilloHeader.FontStyle = FontStyles.Italic
            text.Inlines.Add(LineBreak())
            text.Inlines.Add(civilloHeader)

            if entry.get("civilloDisplay"):
                civilloDisplay = entry.get("civilloDisplay", "")
                slashIndex = civilloDisplay.rfind("/")
                if slashIndex >= 0:
                    civilloFolderText = civilloDisplay[:slashIndex + 1]
                    civilloNameText = civilloDisplay[slashIndex + 1:]
                else:
                    civilloFolderText = ""
                    civilloNameText = civilloDisplay

                civilloFolderRun = Run(civilloFolderText)
                civilloFolderRun.Foreground = SolidColorBrush(Colors.Black)

                civilloNameRun = Run(civilloNameText)
                civilloNameRun.Foreground = SolidColorBrush(Colors.SteelBlue)

                text.Inlines.Add(LineBreak())
                text.Inlines.Add(civilloFolderRun)
                text.Inlines.Add(civilloNameRun)

            if entry.get("civilloCreateNewLayer"):
                newLayerRun = Run("+ new layer: " + (entry.get("civilloNewLayerTitleTemplate") or "{YYMMDD} "))
                newLayerRun.Foreground = SolidColorBrush(Colors.SteelBlue)
                text.Inlines.Add(LineBreak())
                text.Inlines.Add(newLayerRun)

        if entry.get("trimbleConnectTargetName"):
            tcFolderText = "/" + "/".join(entry.get("trimbleConnectFolderPath", []) + [""])
            tcNameText = entry.get("trimbleConnectTargetName", "")
            if entry.get("trimbleConnectTargetIsFolder"):
                tcNameText += " (folder - filename derived at sync time)"

            tcFolderRun = Run(tcFolderText)
            tcFolderRun.Foreground = SolidColorBrush(Colors.Black)

            tcNameRun = Run(tcNameText)
            tcNameRun.Foreground = SolidColorBrush(Colors.SteelBlue)

            tcHeader = Run("Trimble Connect")
            tcHeader.FontWeight = FontWeights.Bold
            tcHeader.FontStyle = FontStyles.Italic
            text.Inlines.Add(LineBreak())
            text.Inlines.Add(tcHeader)
            text.Inlines.Add(LineBreak())
            text.Inlines.Add(tcFolderRun)
            text.Inlines.Add(tcNameRun)

        Grid.SetColumn(text, 1)
        row.Children.Add(text)

        editBtn = Button()
        editBtn.Content = "Edit"
        editBtn.Width = 40
        editBtn.Margin = Thickness(0, 0, 4, 0)
        Grid.SetColumn(editBtn, 2)
        editBtn.Click += lambda s, e, i=index: self.edit_propeller_schedule_clicked(i)
        row.Children.Add(editBtn)

        delBtn = Button()
        delBtn.Content = "X"
        delBtn.Width = 24
        Grid.SetColumn(delBtn, 3)
        delBtn.Click += lambda s, e, i=index: self.delete_propeller_schedule_clicked(i)
        row.Children.Add(delBtn)

        return row

    def propeller_schedule_enabled_toggled(self, index, value):
        if 0 <= index < len(self.propellerSchedules):
            self.propellerSchedules[index]["enabled"] = value
            self.save_propeller_schedules()

    def delete_propeller_schedule_clicked(self, index):
        if 0 <= index < len(self.propellerSchedules):
            del self.propellerSchedules[index]
            self.save_propeller_schedules()
            self.refresh_propeller_schedules_ui()


    # ---------- Propeller org/site/survey selection (drives what the Add/Edit dialog's file list shows) ----------

    def reload_propeller_orgs_clicked(self, sender, e):
        self.propellerError.Content = ""
        self.propOrgCombo.Items.Clear()
        self.propSiteCombo.Items.Clear()
        self.propSurveyCombo.Items.Clear()
        self.propellerSites = []
        self.propellerSelectedFiles = []

        try:
            orgs = self.propellerClient.get_organizations()
        except Exception as ex:
            self.propellerError.Content = str(ex)
            return

        for org in sorted(orgs, key=lambda o: o.get("name", str(o.get("id")))):
            item = ComboBoxItem()
            item.Content = org.get("name", str(org.get("id")))
            item.Tag = org.get("id")
            self.propOrgCombo.Items.Add(item)

        self._select_preferred_propeller(self.propOrgCombo, "SCR_CloudRevise.selectedproporgid")

    def propeller_org_selection_changed(self, sender, e):
        self.propSurveyCombo.Items.Clear()
        self.propSiteCombo.Items.Clear()
        self.propellerSites = []
        self.propellerSelectedFiles = []
        self.propellerError.Content = ""
        self.propSiteFilterBox.Text = ""  # a leftover filter from the previous org wouldn't apply here

        orgItem = self.propOrgCombo.SelectedItem
        if orgItem is None:
            return
        OptionsManager.SetValue("SCR_CloudRevise.selectedproporgid", str(orgItem.Tag))
        # OptionsManager.SetValue only updates the in-memory option store - it's TBC's own controlled
        # shutdown that normally flushes that to Options.bin on disk. A crash or a Task Manager kill of
        # TBC skips that flush entirely, silently reverting this (and window size, etc.) to whatever was
        # last written to disk - which can be an entire session behind. Flushing explicitly here means
        # the org/site choice survives even an unclean TBC exit a moment later.
        OptionsManager.SaveOptions()

        try:
            self.propellerSites = self.propellerClient.get_sites(orgItem.Tag)
        except Exception as ex:
            self.propellerError.Content = str(ex)
            return

        # every ComboBoxItem is built exactly once here, per org selection - filtering afterward (see
        # propeller_site_filter_changed) only ever toggles each item's Visibility rather than touching
        # Items itself, since clearing/rebuilding Items on every filter-box keystroke was the real cause
        # of the filter box feeling slow: Items.Clear() drops the selection to None (firing
        # SelectionChanged), and every rebuilt item is a new instance even for "the same" site, so WPF
        # fires SelectionChanged again for that too - two live get_surveys() candidates per keystroke.
        # Toggling Visibility touches neither Items nor (usually) SelectedItem, so no event fires at all
        # while the current selection stays visible.
        for site in sorted(self.propellerSites, key=lambda s: s.get("name", str(s.get("id")))):
            item = ComboBoxItem()
            item.Content = site.get("name", str(site.get("id")))
            item.Tag = site.get("id")
            self.propSiteCombo.Items.Add(item)

        self._select_preferred_propeller(self.propSiteCombo, "SCR_CloudRevise.selectedpropsiteid")

    def propeller_site_filter_changed(self, sender, e):
        # Visibility toggling itself is immediate on every keystroke (cheap, local, no event) - only the
        # reassignment that happens when the current pick got filtered out is debounced below
        filterText = self.propSiteFilterBox.Text.strip().lower()
        selected = self.propSiteCombo.SelectedItem
        selectedStillVisible = False

        for item in self.propSiteCombo.Items:
            matches = (not filterText) or (filterText in str(item.Content).lower())
            item.Visibility = Visibility.Visible if matches else Visibility.Collapsed
            if matches and item is selected:
                selectedStillVisible = True

        self._propSiteFilterDebounceTimer.Stop()
        if not selectedStillVisible:
            # wait until typing has paused for half a second before actually reassigning the selection -
            # that reassignment is what triggers the real get_surveys() call, so this keeps a fast typist
            # composing a longer search term from firing it for every intermediate site along the way
            self._propSiteFilterDebounceTimer.Start()

    def prop_site_filter_debounce_elapsed(self, sender, e):
        self._propSiteFilterDebounceTimer.Stop()

        selected = self.propSiteCombo.SelectedItem
        if selected is not None and selected.Visibility == Visibility.Visible:
            return  # already resolved in the meantime (e.g. the user picked one manually) - nothing to fix

        firstVisible = None
        for item in self.propSiteCombo.Items:
            if item.Visibility == Visibility.Visible:
                firstVisible = item
                break

        self.propSiteCombo.SelectedItem = firstVisible

    def propeller_site_selection_changed(self, sender, e):
        siteItem = self.propSiteCombo.SelectedItem
        siteId = siteItem.Tag if siteItem is not None else None

        self.refresh_propeller_schedules_ui()  # cheap and local - keep this in sync on every firing below

        # _populate_prop_site_combo() rebuilds propSiteCombo.Items from scratch on every filter-box
        # keystroke, which actually fires SelectionChanged *twice*: once when Items.Clear() drops the
        # selection to None, and again for the (re)selected item - a brand new ComboBoxItem instance even
        # when it's conceptually "the same" site, so WPF raises the event regardless. The first firing
        # used to reset the guard's memory back to None (since None != the real previous id), which let
        # the second firing slip through as if the site had genuinely changed - so it kept hitting
        # get_surveys() on every keystroke despite the guard. Ignoring the transient None firing entirely
        # fixes that: the guard's memory only ever moves between real site ids.
        if siteId is None or siteId == self._lastProcessedPropSiteId:
            return
        self._lastProcessedPropSiteId = siteId

        self.propSurveyCombo.Items.Clear()
        self.propellerSelectedFiles = []
        self.propellerError.Content = ""

        orgItem = self.propOrgCombo.SelectedItem
        if orgItem is None:
            return
        OptionsManager.SetValue("SCR_CloudRevise.selectedpropsiteid", str(siteItem.Tag))
        OptionsManager.SaveOptions()  # see the comment on the same call in propeller_org_selection_changed

        try:
            surveys = self.propellerClient.get_surveys(orgItem.Tag, siteItem.Tag)
        except Exception as ex:
            self.propellerError.Content = str(ex)
            return

        self.propellerSurveys = surveys  # kept for date_captured lookups at run time - see run_propeller_schedules_clicked

        for survey in sorted(surveys, key=lambda s: s.get("date_captured") or s.get("name", str(s.get("id"))), reverse=True):
            item = ComboBoxItem()
            item.Content = survey.get("name", str(survey.get("id")))
            item.Tag = survey.get("id")
            self.propSurveyCombo.Items.Add(item)

        self._select_newest_survey(surveys)

    def _select_newest_survey(self, surveys):
        # always jumps to the newest survey by its real date_captured field (ISO 8601 - confirmed against
        # Propeller's API reference), never a remembered selection - unlike org/site, "whatever I had
        # picked last time" goes stale here the moment a newer survey exists, whether that's after a site
        # change or a fresh startup. Using date_captured rather than parsing a leading date out of the
        # name string means this doesn't depend on every survey actually being named that way.
        newestId, newestDate = None, None
        for survey in surveys:
            dateCaptured = survey.get("date_captured")
            if dateCaptured and (newestDate is None or dateCaptured > newestDate):
                newestId, newestDate = survey.get("id"), dateCaptured

        if newestId is not None:
            for item in self.propSurveyCombo.Items:
                if item.Tag == newestId:
                    self.propSurveyCombo.SelectedItem = item
                    return

        if self.propSurveyCombo.Items.Count > 0:
            self.propSurveyCombo.SelectedIndex = 0

    def propeller_survey_selection_changed(self, sender, e):
        self.propellerSelectedFiles = []
        self.propellerError.Content = ""

        orgItem = self.propOrgCombo.SelectedItem
        siteItem = self.propSiteCombo.SelectedItem
        surveyItem = self.propSurveyCombo.SelectedItem
        if orgItem is None or siteItem is None or surveyItem is None:
            return

        try:
            self.propellerSelectedFiles = self.propellerClient.get_survey_files(orgItem.Tag, siteItem.Tag, surveyItem.Tag)
        except Exception as ex:
            self.propellerError.Content = str(ex)

    def _select_preferred_propeller(self, combo, optionKey):
        # remembers the last-picked org/site/survey across sessions - there's no per-entry "preferred"
        # concept here any more (unlike the Add dialog's civillo/TC pickers), since org/site/survey now
        # live on this window rather than being chosen per sync entry
        preferred = OptionsManager.GetString(optionKey, "") or None
        target = None
        if preferred is not None:
            for item in combo.Items:
                if str(item.Tag) == str(preferred):
                    target = item
                    break

        if target is not None:
            combo.SelectedItem = target
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0

    def _current_propeller_context(self):
        orgItem = self.propOrgCombo.SelectedItem
        siteItem = self.propSiteCombo.SelectedItem
        surveyItem = self.propSurveyCombo.SelectedItem
        if orgItem is None or siteItem is None or surveyItem is None:
            return None

        # self.propellerSites (fetched for the site filter box) already carries each site's full record,
        # coordinate_reference_system_pointer included - no extra API call needed to label files by CRS
        site = next((s for s in self.propellerSites if s.get("id") == siteItem.Tag), None)

        return {
            "orgId": orgItem.Tag,
            "orgName": str(orgItem.Content),
            "siteId": siteItem.Tag,
            "siteName": str(siteItem.Content),
            "siteCrsLabel": format_site_crs_label(site),
            "surveyId": surveyItem.Tag,
            "surveyName": str(surveyItem.Content),
            "files": self.propellerSelectedFiles or [],
            "pdalAvailable": self.pdalAvailable,
            "pdalExePath": self.pdalExePath,
        }


    # ---------- Propeller -> Civillo/Trimble Connect UI actions ----------

    def show_connecting_status(self):
        # SCR_CloudReviseAddPropellerSyncDialog's __init__ synchronously loads Civillo orgs and Trimble
        # Connect regions before the window ever appears, so there's a visible pause with no feedback
        # otherwise - show a message and pump the dispatcher once so it actually paints before we block
        # on that construction. Propeller org/site/survey/files are already loaded on this window itself.
        self.propellerError.Content = "Connecting to Civillo and Trimble Connect..."
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

    def add_propeller_schedule_clicked(self, sender, e):
        propellerContext = self._current_propeller_context()
        if propellerContext is None:
            self.propellerError.Content = "Select a Propeller organization, site, and survey first."
            return

        self.show_connecting_status()
        dlg = SCR_CloudReviseAddPropellerSyncDialog(self.macroFileFolder, self.civilloClient, propellerContext)
        self.propellerError.Content = ""

        dlg.Owner = self
        dlg.ShowDialog()

        if dlg.result is not None:
            entry = dlg.result
            entry["enabled"] = True
            self.propellerSchedules.append(entry)
            self.save_propeller_schedules()
            self.refresh_propeller_schedules_ui()

    def edit_propeller_schedule_clicked(self, index):
        if not (0 <= index < len(self.propellerSchedules)):
            return

        propellerContext = self._current_propeller_context()
        if propellerContext is None:
            self.propellerError.Content = "Select a Propeller organization, site, and survey first."
            return

        existingEntry = self.propellerSchedules[index]

        self.show_connecting_status()
        dlg = SCR_CloudReviseAddPropellerSyncDialog(self.macroFileFolder, self.civilloClient, propellerContext, existingEntry=existingEntry)
        self.propellerError.Content = ""

        dlg.Owner = self
        dlg.ShowDialog()

        if dlg.result is not None:
            enabled = existingEntry.get("enabled", True)
            dlg.result["enabled"] = enabled
            self.propellerSchedules[index] = dlg.result
            self.save_propeller_schedules()
            self.refresh_propeller_schedules_ui()

    def run_propeller_schedules_clicked(self, sender, e):
        # a schedule entry no longer pins a specific survey - it only remembers which site it belongs to
        # (propellerSiteId) and the deliverable's name string (propellerFileRawName), since Propeller uses
        # the same file names for every survey. Running looks that name up in whatever survey is CURRENTLY
        # selected on this window, so the same entry can be re-run against a new survey just by changing
        # the survey dropdown - no need to re-add it.
        siteItem = self.propSiteCombo.SelectedItem
        surveyItem = self.propSurveyCombo.SelectedItem
        if siteItem is None or surveyItem is None:
            MessageBox.Show("Select a Propeller site and survey first.", "SCR_CloudRevise")
            return

        siteId = siteItem.Tag
        siteName = str(siteItem.Content)
        surveyName = str(surveyItem.Content)
        availableFiles = self.propellerSelectedFiles or []

        # for "create additional new layer"'s {YYMMDD} substitution - the survey's own capture date, not
        # today's date, so re-running the same survey's sync later doesn't create a new layer every time
        surveyRecord = next((s for s in self.propellerSurveys if s.get("id") == surveyItem.Tag), None)
        surveyYYMMDD = format_survey_yymmdd(surveyRecord.get("date_captured") if surveyRecord else None)

        enabledEntries = [entry for entry in self.propellerSchedules
                          if entry.get("enabled", True) and entry.get("propellerSiteId") == siteId]
        if len(enabledEntries) == 0:
            MessageBox.Show("No enabled sync entries for the currently selected site.", "SCR_CloudRevise")
            return

        self._begin_cancellable_run()
        results = []
        try:
            for entry in enabledEntries:
                label = entry.get("propellerFileRawName") or entry.get("propellerFileName", "unknown")
                try:
                    matchedFile = next((f for f in availableFiles if f.get("name") == entry.get("propellerFileRawName")), None)
                    if matchedFile is None:
                        results.append(label + " -> SKIPPED (not found in the current survey's file list)")
                        continue
                    results.append(label + " -> " + self.run_single_propeller_sync(entry, siteName, surveyName, matchedFile, surveyYYMMDD))
                except PropellerRunCancelled:
                    results.append(label + " -> CANCELLED")
                    break  # stop processing the remaining entries too, not just this one
                except Exception as ex:
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    message = str(ex) + " (line " + str(exc_tb.tb_lineno) + ")"
                    results.append(label + " -> FAILED: " + message)
                    try:
                        if not os.path.isdir(r"C:\temp"):
                            os.makedirs(r"C:\temp")
                        with open(r"C:\temp\propeller_run.log", "a") as f:
                            f.write(str(datetime.now()) + " " + label + "\n" + message + "\n\n")
                    except Exception:
                        pass
                finally:
                    self.set_run_progress(None)
        finally:
            self._end_cancellable_run()

        MessageBox.Show("\n".join(results), "SCR_CloudRevise - Run Results")

    def _begin_cancellable_run(self):
        self._propellerCancelRequested = False
        self.cancelPropellerRunBtn.IsEnabled = True
        self.runPropellerSchedulesBtn.IsEnabled = False
        self.addPropellerScheduleBtn.IsEnabled = False

    def _end_cancellable_run(self):
        self.cancelPropellerRunBtn.IsEnabled = False
        self.runPropellerSchedulesBtn.IsEnabled = True
        self.addPropellerScheduleBtn.IsEnabled = True

        if self._closeRequestedDuringRun:
            # the window's own Closing handler deferred this close until the run (now finished
            # unwinding) was done with it - re-issue it now; SaveOptions will see no run in progress
            # this time and let it proceed normally
            self._closeRequestedDuringRun = False
            self.Close()

    def cancel_propeller_run_clicked(self, sender, e):
        # cooperative cancellation only - there's no background thread here, the run is still executing
        # synchronously on this same UI thread between set_run_progress() calls (see the DispatcherPriority
        # note there for how this click even gets a chance to run at all). Setting the flag doesn't stop
        # anything by itself; the next set_run_progress() call downstream (resample progress, download
        # progress, or the next schedule entry starting) is what actually raises to unwind out cleanly.
        self._propellerCancelRequested = True
        self.cancelPropellerRunBtn.IsEnabled = False
        self.propellerProgressStatus.Text = self.propellerProgressStatus.Text + " (cancelling...)"

    def set_run_progress(self, statusText, percent=None):
        # a visible progress bar on the window itself, rather than TBC's own status-bar progress control
        # (still used separately for SaveFileRemotely's native IProgressBarControl parameter, where it's a
        # real API requirement, not a stylistic choice) - statusText=None hides the panel again. Pumps the
        # dispatcher once so the update actually paints during a long synchronous run instead of only
        # showing up once everything's already finished.
        if statusText:
            self.propellerProgressPanel.Visibility = Visibility.Visible
            self.propellerProgressStatus.Text = statusText
            self.propellerProgressBar.Value = percent if percent is not None else 0
        else:
            self.propellerProgressPanel.Visibility = Visibility.Collapsed
            self.propellerProgressBar.Value = 0

        # DispatcherPriority.Input (not Render) - WPF's priority order puts mouse clicks at Input, which
        # is BELOW Render, so a Render-priority pump (the old value here) never lets the Cancel button's
        # own Click event through while a run is still going - it'd only fire once the whole synchronous
        # run already finished, making Cancel a no-op. Input is low enough to drain the click along with
        # everything above it (Render included), so the repaint this call exists for still happens too.
        self.Dispatcher.Invoke(DispatcherPriority.Input, Action(lambda: None))

        # only the "showing progress" branch checks for cancellation - never the statusText=None branch,
        # which every caller (including a just-cancelled one, from its own finally block) uses to hide the
        # panel again on the way out; raising there too would replace a clean single cancellation with a
        # second one escaping from inside that cleanup itself
        if statusText and self._propellerCancelRequested:
            raise PropellerRunCancelled()

    def run_single_propeller_sync(self, entry, siteName, surveyName, matchedFile, surveyYYMMDD=None):
        # matchedFile is a live file dict looked up (by name) from the currently selected survey's file
        # list - its url/format/size are always fresh, unlike anything that would have been stored on the
        # entry itself (which only remembers the file's name string, not an id that could go stale or tie
        # it to one specific survey)
        fileInfo = {
            "url": matchedFile.get("url"),
            "name": matchedFile.get("name", "file"),
            "format": matchedFile.get("format") or "tiff",
        }

        def reportDownloadProgress(bytesWritten, totalBytes):
            if totalBytes > 0:
                self.set_run_progress("Downloading " + fileInfo["name"] + "...", int(bytesWritten * 100 / totalBytes))

        self.set_run_progress("Downloading " + fileInfo["name"] + "...", 0)
        sourcePath, downloadCached = self.propellerDownloadCache.get_or_download(self.propellerClient, siteName, surveyName, fileInfo, reportDownloadProgress)
        if downloadCached:
            self.set_run_progress(fileInfo["name"] + " found in cache - skipping download", 100)

        isLaz = fileInfo["format"] == "laz"

        # only meaningful for isLaz: the real filename Propeller's web UI shows (point count included,
        # when named that way) lives in the download URL itself, not the API's own short "name" field -
        # see extract_filename_from_propeller_url. Used only as a speed hint for pre-sizing arrays, never
        # trusted over what the file actually contains.
        expectedPointCount = None
        if isLaz:
            realFilename = extract_filename_from_propeller_url(matchedFile.get("url") or "")
            if realFilename:
                expectedPointCount = parse_point_count_from_filename(realFilename)

        pushed = []

        civHasTarget = entry.get("civilloLayerId") is not None or entry.get("civilloCreateNewLayer")
        if entry.get("civilloEnabled", True) and civHasTarget:
            if isLaz:
                civMaxSizeGb = entry.get("civilloMaxSizeGb")
                civPath, civWorldfile = self._prepare_platform_laz_surface(sourcePath, siteName, surveyName, "Civillo", expectedPointCount,
                                                                            entry.get("civilloLazResampleMethod", LAZ_RESAMPLE_METHOD),
                                                                            entry.get("civilloLazResampleSpacing", LAZ_RESAMPLE_SPACING),
                                                                            maxSizeBytes=civMaxSizeGb * (1024 ** 3) if civMaxSizeGb else None)
            else:
                civPath, civWorldfile = self._prepare_platform_file(sourcePath, entry, "civilloPixelSizeMeters", "civilloOutputFormat", "civilloMaxSizeGb", siteName, surveyName, "Civillo")
            files = [(civPath, os.path.basename(civPath))]
            if civWorldfile is not None:
                files.append((civWorldfile, os.path.basename(civWorldfile)))
            if not isLaz:
                files = zip_files_for_civillo_ortho(files)

            civilloUploadStartTime = timer()

            def formatDurationSeconds(totalSeconds):
                # "Xm Ys" (or just "Ys" under a minute) - a rough upload ETA, not meant to be precise to
                # the second, just a sense of how much longer this upload will take
                totalSeconds = int(round(totalSeconds))
                if totalSeconds < 60:
                    return str(totalSeconds) + "s"
                minutes, seconds = divmod(totalSeconds, 60)
                return str(minutes) + "m " + str(seconds) + "s"

            def reportCivilloUploadProgress(bytesSent, totalBytes):
                if totalBytes > 0:
                    # IronPython ints are real .NET integers - ToString("N0") gives thousands separators
                    # (str.format's "{:,}" doesn't work the same way in IronPython, hence going straight to .NET)
                    statusText = "Uploading to Civillo... (" + bytesSent.ToString("N0") + " / " + totalBytes.ToString("N0") + " bytes)"

                    elapsed = timer() - civilloUploadStartTime
                    # skip the estimate for the first fraction of a second - too little data yet for
                    # bytesSent/elapsed to mean anything, so the earliest ETA would just be a wild guess
                    if bytesSent > 0 and elapsed > 0.5:
                        bytesPerSecond = bytesSent / elapsed
                        remainingSeconds = (totalBytes - bytesSent) / bytesPerSecond
                        statusText += " - ETA " + formatDurationSeconds(remainingSeconds)

                    self.set_run_progress(statusText, int(bytesSent * 100 / totalBytes))

            if entry.get("civilloLayerId") is not None:
                self.set_run_progress("Uploading to Civillo...", 0)
                push_to_civillo(self.civilloClient, entry["civilloOrgNickname"], entry["civilloProjectId"], entry["civilloLayerId"], files, reportCivilloUploadProgress)
                pushed.append("Civillo")

            if entry.get("civilloCreateNewLayer"):
                titleTemplate = entry.get("civilloNewLayerTitleTemplate") or "{YYMMDD} "
                newTitle = titleTemplate.replace("{YYMMDD}", surveyYYMMDD) if surveyYYMMDD else titleTemplate
                self.set_run_progress("Creating new Civillo layer...", 0)
                push_to_civillo_create_new(self.civilloClient, entry["civilloOrgNickname"], entry["civilloProjectId"], newTitle, files, reportCivilloUploadProgress)
                pushed.append("Civillo (new layer)")

        if entry.get("trimbleConnectEnabled", True) and entry.get("trimbleConnectTargetName") is not None:
            tcTarget = self._resolve_tc_upload_folder(entry)
            if isLaz:
                tcMaxSizeGb = entry.get("trimbleConnectMaxSizeGb")
                tcPath, tcWorldfile = self._prepare_platform_laz_surface(sourcePath, siteName, surveyName, "TrimbleConnect", expectedPointCount,
                                                                          entry.get("trimbleConnectLazResampleMethod", LAZ_RESAMPLE_METHOD),
                                                                          entry.get("trimbleConnectLazResampleSpacing", LAZ_RESAMPLE_SPACING),
                                                                          maxSizeBytes=tcMaxSizeGb * (1024 ** 3) if tcMaxSizeGb else None)
            else:
                tcPath, tcWorldfile = self._prepare_platform_file(sourcePath, entry, "trimbleConnectPixelSizeMeters", "trimbleConnectOutputFormat", "trimbleConnectMaxSizeGb", siteName, surveyName, "TrimbleConnect")
            self.set_run_progress("Uploading to Trimble Connect...")
            self.trimbleConnectClient.save_file_remotely(tcPath, tcTarget)
            if tcWorldfile is not None:
                self.trimbleConnectClient.save_file_remotely(tcWorldfile, tcTarget)
            pushed.append("Trimble Connect")

        self.set_run_progress(None)

        if not pushed:
            return "skipped (no enabled target)"
        return "OK -> " + ", ".join(pushed)

    def _prepare_platform_laz_surface(self, sourcePath, siteName, surveyName, platformTag, expectedPointCount, resampleMethod, resampleSpacing, maxSizeBytes=None):
        # same cache subfolder shape _prepare_platform_file uses for a resampled GeoTIFF/JPEG
        # (<cache>/<site>/<survey>/<platform>/)
        subfolder = os.path.join(self.get_propeller_cache_folder(), sanitize_filename(siteName), sanitize_filename(surveyName), platformTag)
        if not os.path.isdir(subfolder):
            os.makedirs(subfolder)

        if not resampleSpacing:
            # nothing to grow if the size limit isn't met - there's no spacing dial to turn, so the
            # limit simply isn't enforceable here (in practice this doesn't happen: the Add dialog always
            # has a numeric spacing for a LAZ entry, this is just a defensive fallback)
            maxSizeBytes = None

        # Every attempt resamples from the original sourcePath (never a previous attempt's own already-
        # thinned LandXML) - same reasoning as _resample_to_size_limit's own loop for orthophotos:
        # resampling an already-resampled result would compound thinning for no benefit when the
        # original LAZ is right there. Each attempt's spacing is baked into its filename, so a coarser
        # retry never collides with an earlier one's - earlier attempts are removed below once the one
        # that actually gets uploaded is known.
        maxAttempts = 6
        outputPath = None
        unusedAttempts = []
        for attempt in range(maxAttempts):
            # "%g" rather than repr() - repr() on a float can emit up to ~17 raw digits (harmless for the
            # actual resampling, but produces an unreadably long filename, e.g. "..._1_2672257297509337.xml"
            # for an estimated retry spacing); the round() above already keeps every value that reaches
            # here at 3 decimals or fewer, so "%g" just formats that cleanly without reintroducing noise
            suffix = "_" + resampleMethod + "_" + ("%g" % resampleSpacing).replace(".", "_") if resampleSpacing else ""
            outputPath = os.path.join(subfolder, sanitize_filename(surveyName) + suffix + ".xml")

            # this exact spacing/method was already resampled (by a previous run, or an earlier attempt
            # this same run) and is still sitting in the cache - skip straight to uploading it rather than
            # re-triangulating an identical result
            if os.path.isfile(outputPath):
                self.set_run_progress(os.path.basename(outputPath) + " found in cache - skipping resample", 100)
            else:
                # progressCallback's contract (statusText, percent=None) matches set_run_progress's own
                # signature exactly - no wrapper needed
                laz_points_to_landxml_surface(sourcePath, outputPath, surfaceName=sanitize_filename(surveyName), progressCallback=self.set_run_progress,
                                               resampleSpacing=resampleSpacing, resampleMethod=resampleMethod, expectedPointCount=expectedPointCount,
                                               pdalExePath=self.pdalExePath)

            if maxSizeBytes is None:
                return outputPath, None  # a LandXML file is self-contained - no separate worldfile needed

            actualSizeBytes = os.path.getsize(outputPath)
            if actualSizeBytes <= maxSizeBytes:
                self._remove_files(unusedAttempts)
                return outputPath, None

            unusedAttempts.append(outputPath)

            if attempt == maxAttempts - 1:
                # never got under the limit - the last (still oversized) attempt gets thrown away too
                # rather than uploaded anyway, same as _resample_to_size_limit's own last-attempt handling
                self._remove_files(unusedAttempts)
                raise Exception(
                    "Could not resample the " + platformTag + " surface under the " + ("%.3f" % (maxSizeBytes / (1024.0 ** 3))) +
                    " GB limit after " + str(maxAttempts) + " attempts (best attempt: " +
                    ("%.3f" % (actualSizeBytes / (1024.0 ** 3))) + " GB at " + ("%.3f" % resampleSpacing) + "m spacing)")

            # LandXML size scales roughly with kept point count, which for a roughly uniform-density cloud
            # scales ~1/spacing^2 regardless of resample method (voxel/grid/octree/pdal all target a
            # roughly uniform output density at a given spacing) - same sqrt(ratio) estimate
            # _resample_to_size_limit uses for pixel size, with the same small safety margin
            ratio = float(actualSizeBytes) / float(maxSizeBytes)
            # rounded to 3 decimals - both to keep the estimate at a sensible precision for a spacing
            # value (nobody needs mm-of-a-mm resolution here) and so the filename below stays short and
            # readable instead of carrying ~17 raw floating-point digits (e.g. "..._1_2672257297509337.xml")
            nextSpacing = round(resampleSpacing * math.sqrt(ratio) * 1.03, 3)
            if round(nextSpacing, 2) <= round(resampleSpacing, 2):
                # the estimate didn't move the spacing enough to change anything meaningful - force a step
                # forward so the next attempt actually differs (and gets a distinct filename) from this one
                nextSpacing = round(resampleSpacing + 0.01, 3)
            resampleSpacing = nextSpacing

        self._remove_files(unusedAttempts)
        return outputPath, None

    def _prepare_platform_file(self, sourcePath, entry, pixelKey, formatKey, maxSizeKey, siteName, surveyName, platformTag):
        configuredPixelSizeMeters = entry.get(pixelKey)
        outputFormat = entry.get(formatKey) or "TIFF"
        maxSizeGb = entry.get(maxSizeKey)
        # "GB" here means GiB (1024^3) to match what Windows Explorer's own file size column shows under
        # that label, since that's what a user is most likely comparing this setting against
        maxSizeBytes = maxSizeGb * (1024 ** 3) if maxSizeGb else None

        # Propeller's own deliverable is always a GeoTIFF, so a configured TIFF output with no resample
        # pixel size is a genuine no-op format-wise - the only reason to touch the file at all then is the
        # max-GB limit. A configured JPEG output is never a no-op, even with no pixel size: it still needs
        # an actual GDAL pass to re-encode as JPEG, it just shouldn't resample the resolution while at it.
        isNativeFormat = outputFormat.upper() == "TIFF"

        if configuredPixelSizeMeters is None:
            if isNativeFormat and (maxSizeBytes is None or os.path.getsize(sourcePath) <= maxSizeBytes):
                # nothing to do at all - push the original straight from the download cache as-is: no
                # GDAL pass, and no copy into this platform's own subfolder either
                return sourcePath, None
            # either a JPEG re-encode (always needs a GDAL pass) or an oversized TIFF (needs the max-GB
            # retry loop) with no requested pixel size to start from - fall back to the source's own
            # native ground sample distance so the first attempt doesn't resample anything, only the
            # usual retry loop below (if it even runs) grows it from there
            pixelSizeMeters = get_geotiff_pixel_size_meters(sourcePath)
        else:
            pixelSizeMeters = configuredPixelSizeMeters

        ext = "jpg" if outputFormat.upper() == "JPEG" else "tif"

        subfolder = os.path.join(self.get_propeller_cache_folder(), sanitize_filename(siteName), sanitize_filename(surveyName), platformTag)
        if not os.path.isdir(subfolder):
            os.makedirs(subfolder)

        outputPath, worldfilePath, finalPixelSizeMeters = self._resample_to_size_limit(sourcePath, subfolder, surveyName, pixelSizeMeters, outputFormat, ext, platformTag, maxSizeBytes)

        # persist a pixel size that grew (JPEG's dimension clamp, or the max-GB retry loop) back into the
        # schedule entry, so next run starts from the size that actually worked instead of re-discovering
        # it through the same handful of oversized attempts every single time - only when a pixel size
        # was actually configured, though: landing here via the native-GSD fallback above shouldn't
        # silently turn resampling on for every future run of an entry that never asked for it
        if configuredPixelSizeMeters is not None and finalPixelSizeMeters != configuredPixelSizeMeters:
            entry[pixelKey] = finalPixelSizeMeters
            self.save_propeller_schedules()

        return outputPath, worldfilePath

    def _resample_to_size_limit(self, sourcePath, subfolder, surveyName, pixelSizeMeters, outputFormat, ext, platformTag, maxSizeBytes):
        if outputFormat.upper() == "JPEG":
            # JPEG hard-caps each dimension at JPEG_MAX_DIMENSION_PX regardless of the max-size-GB setting
            # below - a pixel size fine enough to blow past that (e.g. a few mm/pixel over a large site)
            # fails outright with a raw libjpeg error rather than just producing an oversized file, so
            # clamp it up front here rather than let the translate attempt hit that wall. Only ever needs
            # doing once: every later attempt in the loop below only grows the pixel size further, so once
            # this floor is satisfied it stays satisfied.
            extentXMeters, extentYMeters = get_geotiff_extent_meters(sourcePath)
            minPixelSizeMeters = max(extentXMeters, extentYMeters) / (JPEG_MAX_DIMENSION_PX - 1) * 1.001
            if pixelSizeMeters < minPixelSizeMeters:
                pixelSizeMeters = minPixelSizeMeters

        # Every attempt resamples from the original sourcePath, never from a previous attempt's own
        # (already-downsampled) output - repeatedly resampling a resample would compound resampling error
        # for no benefit, when the original is right there. Each attempt's pixel size is baked into its
        # filename (derive_output_filename), so a coarser retry never collides with an earlier one's -
        # earlier attempts are removed below once the one that actually gets uploaded is known, so an
        # oversized intermediate resample never sits around unused on disk.
        maxAttempts = 6
        outputPath = worldfilePath = None
        unusedAttempts = []
        for attempt in range(maxAttempts):
            outputPath = os.path.join(subfolder, derive_output_filename(surveyName, pixelSizeMeters, ext))
            expectedWorldfilePath = final_worldfile_path_for(outputPath, ext)

            # this exact pixel size/format was already resampled (by a previous run, or an earlier attempt
            # this same run) and is still sitting in the cache - skip straight to uploading it rather than
            # re-running GDAL for an identical result
            if os.path.isfile(outputPath) and (expectedWorldfilePath is None or os.path.isfile(expectedWorldfilePath)):
                worldfilePath = expectedWorldfilePath
                self.set_run_progress(os.path.basename(outputPath) + " found in cache - skipping resample", 100)
            else:
                worldfilePath = worldfile_path_for(outputPath)

                def reportGdalProgress(percent):
                    self.set_run_progress("Resampling for " + platformTag + "...", percent)

                self.set_run_progress("Resampling for " + platformTag + "...", 0)
                convert_geotiff_resolution(sourcePath, outputPath, pixelSizeMeters, outputFormat, reportGdalProgress)
                # GDAL_PAM_ENABLED=NO (set in convert_geotiff_resolution) should already stop this from being
                # written at all, but delete it unconditionally either way - we don't use it for any attempt,
                # winning or not, so there's no case where it's worth keeping around
                self._remove_files([outputPath + ".aux.xml"])
                worldfilePath = rename_worldfile_for_format(worldfilePath, ext)

            if maxSizeBytes is None:
                return outputPath, worldfilePath, pixelSizeMeters

            actualSizeBytes = os.path.getsize(outputPath)
            if actualSizeBytes <= maxSizeBytes:
                self._remove_files(unusedAttempts)
                return outputPath, worldfilePath, pixelSizeMeters

            unusedAttempts.append(outputPath)
            unusedAttempts.append(worldfilePath)

            if attempt == maxAttempts - 1:
                # never got under the limit - the last (still oversized) attempt gets thrown away too
                # rather than uploaded anyway, same as every other attempt that didn't make the cut
                self._remove_files(unusedAttempts)
                raise Exception(
                    "Could not resample for " + platformTag + " under the " + ("%.3f" % (maxSizeBytes / (1024.0 ** 3))) +
                    " GB limit after " + str(maxAttempts) + " attempts (best attempt: " +
                    ("%.3f" % (actualSizeBytes / (1024.0 ** 3))) + " GB at " + ("%.3f" % pixelSizeMeters) + " m/pixel)")

            # output size scales roughly with pixel count, i.e. ~1/pixelSizeMeters^2 for a fixed ground
            # extent, so growing the pixel size by sqrt(actual/target) should land close to the target on
            # the next pass - a 3% safety margin covers the small amount of extra shrinkage compression
            # variance (especially JPEG's quality-driven ratio) tends to need beyond that estimate
            ratio = float(actualSizeBytes) / float(maxSizeBytes)
            nextPixelSizeMeters = pixelSizeMeters * math.sqrt(ratio) * 1.03
            if int(round(nextPixelSizeMeters * 100)) <= int(round(pixelSizeMeters * 100)):
                # the estimate rounded back into the same cm bucket as derive_output_filename uses -
                # force a step forward so the next attempt gets a distinct filename and actually shrinks
                nextPixelSizeMeters = pixelSizeMeters + 0.01
            pixelSizeMeters = nextPixelSizeMeters

        self._remove_files(unusedAttempts)
        return outputPath, worldfilePath, pixelSizeMeters

    def _remove_files(self, paths):
        for path in paths:
            if path and os.path.isfile(path):
                try:
                    os.remove(path)
                except Exception:
                    pass  # leftover from a resample attempt that didn't get uploaded - not worth failing the run over

    def _resolve_tc_upload_folder(self, entry):
        # walks region -> project -> folder path (all stored as names on the schedule entry, since Trimble
        # Connect gives us no stable ids to persist instead - region.ID is always None) to find the actual
        # folder object SaveFileRemotely needs
        regions = self.trimbleConnectClient.get_regions()
        region = next((r for r in regions if str(r.FileName) == entry.get("trimbleConnectRegionName")), None)
        if region is None:
            raise Exception("Trimble Connect region not found: " + str(entry.get("trimbleConnectRegionName")))

        projects = self.trimbleConnectClient.get_children(region)
        project = next((p for p in projects if str(p.FileName) == entry.get("trimbleConnectProjectName")), None)
        if project is None:
            raise Exception("Trimble Connect project not found: " + str(entry.get("trimbleConnectProjectName")))

        uploadFolderPath = entry.get("trimbleConnectUploadFolderPath") or []
        # self-heal entries saved before this bug was fixed: a first segment that just repeats the
        # project's own name means "topmost folder" was selected and got mistakenly recorded as a
        # subfolder to descend into, rather than an empty path meaning "the project root itself"
        if uploadFolderPath and uploadFolderPath[0] == str(project.FileName):
            uploadFolderPath = uploadFolderPath[1:]

        current = project
        for folderName in uploadFolderPath:
            children = self.trimbleConnectClient.get_children(current)
            match = next((c for c in children if str(c.FileName) == folderName), None)
            if match is None:
                availableNames = [str(c.FileName) for c in children]
                raise Exception("Trimble Connect folder '" + folderName + "' not found under '" + str(getattr(current, "FileName", "?")) + "' - available: " + str(availableNames))
            current = match

        return current


    # ---------- UI actions ----------

    def help_clicked(self, sender, e):
        webbrowser.open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\MacroHelp\MacroHelp.htm#SCR_CloudRevise")

    def reload_orgs_clicked(self, sender, e):
        configPath = self.get_config_path()
        credentials = self.load_or_create_config(configPath)
        self.civilloConfig = credentials["civillo"]
        self.propellerConfig = credentials["propeller"]
        self.civilloClient = CivilloClient(self.civilloConfig)
        self.propellerClient = PropellerClient(self.propellerConfig)

        if self.is_placeholder_config(self.civilloConfig):
            self.error.Content = "Edit the config file above with your Civillo API key/secret, then click 'Reload Organizations'."
            return

        self.load_organizations()

    def load_organizations(self):
        self.error.Content = ""
        self.orgCombo.Items.Clear()
        self.projectCombo.Items.Clear()

        try:
            orgs = self.civilloClient.get("/applications")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for org in orgs:
            item = ComboBoxItem()
            item.Content = org["name"]
            item.Tag = org["nickname"]
            self.orgCombo.Items.Add(item)

        savedNickname, savedProjectId = self.get_saved_selection()
        target = None
        for item in self.orgCombo.Items:
            if item.Tag == savedNickname:
                target = item
                break

        if target is not None:
            self.orgCombo.SelectedItem = target
        elif self.orgCombo.Items.Count > 0:
            self.orgCombo.SelectedIndex = 0

    def org_selection_changed(self, sender, e):
        self.projectCombo.Items.Clear()
        self.error.Content = ""

        selected = self.orgCombo.SelectedItem
        if selected is None:
            return

        nickname = selected.Tag

        try:
            projects = self.civilloClient.get("/" + nickname + "/projects")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for proj in projects:
            item = ComboBoxItem()
            item.Content = proj["name"]
            item.Tag = proj["id"]
            self.projectCombo.Items.Add(item)

        savedNickname, savedProjectId = self.get_saved_selection()
        target = None
        if nickname == savedNickname:
            for item in self.projectCombo.Items:
                if str(item.Tag) == savedProjectId:
                    target = item
                    break

        if target is not None:
            self.projectCombo.SelectedItem = target
        elif self.projectCombo.Items.Count > 0:
            self.projectCombo.SelectedIndex = 0

    def project_selection_changed(self, sender, e):
        orgItem = self.orgCombo.SelectedItem
        projectItem = self.projectCombo.SelectedItem

        if orgItem is not None and projectItem is not None:
            self.save_selection(orgItem.Tag, projectItem.Tag)


    # ---------- schedules UI ----------

    def refresh_schedules_ui(self):
        self.schedulesPanel.Children.Clear()

        # group rows by org/project (not by insertion order) so an "Org - Site" header always
        # sits above every contiguous block of that org/project's entries, however they were added
        indices = list(range(len(self.schedules)))
        indices.sort(key=lambda i: self.schedule_group_key(self.schedules[i]))

        lastGroupKey = None
        for i in indices:
            entry = self.schedules[i]
            groupKey = self.schedule_group_key(entry)
            if groupKey != lastGroupKey:
                self.schedulesPanel.Children.Add(self.build_schedule_group_header(entry))
                lastGroupKey = groupKey
            self.schedulesPanel.Children.Add(self.build_schedule_row(i))

    def schedule_group_key(self, entry):
        return (entry.get("orgName") or entry.get("orgNickname") or "", entry.get("projectName") or str(entry.get("projectId") or ""))

    def build_schedule_group_header(self, entry):
        orgLabel = entry.get("orgName") or entry.get("orgNickname") or "Unknown Org"
        siteLabel = entry.get("projectName") or str(entry.get("projectId") or "Unknown Site")

        header = TextBlock()
        header.Text = orgLabel + " - " + siteLabel
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 8, 0, 4)
        return header

    def build_schedule_row(self, index):
        entry = self.schedules[index]
    
        row = Grid()
        row.Margin = Thickness(0, 0, 0, 8)
    
        colCheck = ColumnDefinition()
        colCheck.Width = GridLength(1, GridUnitType.Auto)
    
        colText = ColumnDefinition()
        colText.Width = GridLength(1, GridUnitType.Star)
    
        colEdit = ColumnDefinition()
        colEdit.Width = GridLength(1, GridUnitType.Auto)
    
        colDelete = ColumnDefinition()
        colDelete.Width = GridLength(1, GridUnitType.Auto)
    
        row.ColumnDefinitions.Add(colCheck)
        row.ColumnDefinitions.Add(colText)
        row.ColumnDefinitions.Add(colEdit)
        row.ColumnDefinitions.Add(colDelete)
    
        chk = CheckBox()
        chk.IsChecked = bool(entry.get("enabled", True))
        Grid.SetColumn(chk, 0)
        chk.Checked += lambda s, e, i=index: self.schedule_enabled_toggled(i, True)
        chk.Unchecked += lambda s, e, i=index: self.schedule_enabled_toggled(i, False)
        row.Children.Add(chk)
    
        # Local paths
        local1Run = Run(entry.get("localPath1", ""))
        local1Run.Foreground = SolidColorBrush(Colors.DarkGreen)
    
        local2Run = Run(entry.get("localPath2", ""))
        local2Run.Foreground = SolidColorBrush(Colors.Brown)
    
        # Split Civillo path into folder and name
        civilloPath = entry.get("civilloPath", "")
        slashIndex = civilloPath.rfind("/")
    
        if slashIndex >= 0:
            civilloFolderText = civilloPath[:slashIndex + 1]
            civilloNameText = civilloPath[slashIndex + 1:]
        else:
            civilloFolderText = ""
            civilloNameText = civilloPath
    
        civilloFolderRun = Run(civilloFolderText)
        civilloFolderRun.Foreground = SolidColorBrush(Colors.Black)
    
        civilloNameRun = Run(civilloNameText)
        civilloNameRun.Foreground = SolidColorBrush(Colors.SteelBlue)
    
        # Text block
        text = TextBlock()
        text.TextWrapping = TextWrapping.Wrap
        text.Margin = Thickness(6, 0, 6, 0)
    
        # Local heading
        localHeader = Run("Local")
        localHeader.FontWeight = FontWeights.Bold
        localHeader.FontStyle = FontStyles.Italic
        text.Inlines.Add(localHeader)
        text.Inlines.Add(LineBreak())
    
        # Local Path 1
        text.Inlines.Add(local1Run)
    
        # Local Path 2 (optional)
        if entry.get("localPath2", ""):
            text.Inlines.Add(LineBreak())
            text.Inlines.Add(local2Run)
    
        # Spacer
        text.Inlines.Add(LineBreak())
    
        # Civillo heading
        civilloHeader = Run("Civillo")
        civilloHeader.FontWeight = FontWeights.Bold
        civilloHeader.FontStyle = FontStyles.Italic
        text.Inlines.Add(civilloHeader)
        text.Inlines.Add(LineBreak())
    
        # Civillo path
        text.Inlines.Add(civilloFolderRun)
        text.Inlines.Add(civilloNameRun)
    
        Grid.SetColumn(text, 1)
        row.Children.Add(text)
    
        editBtn = Button()
        editBtn.Content = "Edit"
        editBtn.Width = 40
        editBtn.Margin = Thickness(0, 0, 4, 0)
        Grid.SetColumn(editBtn, 2)
        editBtn.Click += lambda s, e, i=index: self.edit_schedule_clicked(i)
        row.Children.Add(editBtn)
    
        delBtn = Button()
        delBtn.Content = "X"
        delBtn.Width = 24
        Grid.SetColumn(delBtn, 3)
        delBtn.Click += lambda s, e, i=index: self.delete_schedule_clicked(i)
        row.Children.Add(delBtn)
    
        return row

    def schedule_enabled_toggled(self, index, value):
        if 0 <= index < len(self.schedules):
            self.schedules[index]["enabled"] = value
            self.save_schedules()

    def delete_schedule_clicked(self, index):
        if 0 <= index < len(self.schedules):
            del self.schedules[index]
            self.save_schedules()
            self.refresh_schedules_ui()

    def add_schedule_clicked(self, sender, e):
        orgItem = self.orgCombo.SelectedItem
        projectItem = self.projectCombo.SelectedItem
        nickname = orgItem.Tag if orgItem is not None else None
        projectId = projectItem.Tag if projectItem is not None else None

        dlg = SCR_CloudReviseAddSyncDialog(self.macroFileFolder, self.civilloClient, nickname, projectId)
        dlg.Owner = self
        dlg.ShowDialog()

        if dlg.result is not None:
            entry = {
                "enabled": True,
                "localPath1": dlg.result["localPath1"],
                "localPath2": dlg.result["localPath2"],
                "civilloPath": dlg.result["civilloPath"],
                "civilloLayerId": dlg.result["civilloLayerId"],
                "orgNickname": dlg.result["orgNickname"],
                "projectId": dlg.result["projectId"],
                "orgName": str(orgItem.Content) if orgItem is not None else "",
                "projectName": str(projectItem.Content) if projectItem is not None else "",
            }
            self.schedules.append(entry)
            self.save_schedules()
            self.refresh_schedules_ui()

    def edit_schedule_clicked(self, index):
        if not (0 <= index < len(self.schedules)):
            return

        existingEntry = self.schedules[index]

        dlg = SCR_CloudReviseAddSyncDialog(
            self.macroFileFolder, self.civilloClient,
            existingEntry.get("orgNickname"), existingEntry.get("projectId"),
            existingEntry=existingEntry
        )
        dlg.Owner = self
        dlg.ShowDialog()

        if dlg.result is not None:
            existingEntry["localPath1"] = dlg.result["localPath1"]
            existingEntry["localPath2"] = dlg.result["localPath2"]
            existingEntry["civilloPath"] = dlg.result["civilloPath"]
            existingEntry["civilloLayerId"] = dlg.result["civilloLayerId"]
            existingEntry["orgNickname"] = dlg.result["orgNickname"]
            existingEntry["projectId"] = dlg.result["projectId"]
            existingEntry.pop("lastSyncedHash", None)  # force the next run, even if file 1 is unchanged but file 2 was added/removed
            self.save_schedules()
            self.refresh_schedules_ui()

    def run_schedules_clicked(self, sender, e):
        self.error.Content = ""
        results = []

        for entry in self.schedules:
            if not entry.get("enabled", True):
                continue
            label = os.path.basename(entry.get("localPath1", ""))
            try:
                status = self.run_single_sync(entry)
                if status == "skipped":
                    results.append(label + " -> skipped (unchanged)")
                else:
                    results.append(label + " -> OK")
            except Exception as ex:
                results.append(label + " -> FAILED: " + str(ex))

        if len(results) == 0:
            MessageBox.Show("No enabled sync entries to run.", "SCR_CloudRevise")
        else:
            MessageBox.Show("\n".join(results), "SCR_CloudRevise - Run Results")

    def compute_file_hash(self, path):
        hasher = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def run_single_sync(self, entry):
        localPath1 = entry.get("localPath1", "")
        localPath2 = entry.get("localPath2", "")

        if not os.path.isfile(localPath1):
            raise Exception("Local file not found: " + localPath1)
        if localPath2 and not os.path.isfile(localPath2):
            raise Exception("Local file not found: " + localPath2)

        currentHash = self.compute_file_hash(localPath1)
        lastHash = entry.get("lastSyncedHash")

        if lastHash is not None and lastHash == currentHash:
            return "skipped"

        # Civillo's auto-detection of the source SRS is currently unavailable server-side, so we
        # must pass the project's own default projection explicitly instead of relying on srid=-1 (auto).
        projectInfo = self.civilloClient.get("/" + entry["orgNickname"] + "/projects/" + str(entry["projectId"]))
        srid = projectInfo.get("defaultProjection", -1)

        # local file 2 (e.g. a linestyle file) is optional - only include it if the user supplied one
        files = [(localPath1, os.path.basename(localPath1))]
        if localPath2:
            files.append((localPath2, os.path.basename(localPath2)))

        body = {
            "mode": 1,  # 1 = revise an existing layer
            "fileNames": [name for _, name in files],
            "replaceLayerId": entry["civilloLayerId"],
            "fileLastModifieds": [int(os.path.getmtime(path)) for path, _ in files],
            "srid": srid,
        }

        initResp = self.civilloClient.post_json(
            "/" + entry["orgNickname"] + "/projects/" + str(entry["projectId"]) + "/layers",
            body
        )

        self.civilloClient.upload_files(
            initResp["processPath"], initResp["token"], initResp["application"], initResp["job"],
            files
        )

        entry["lastSyncedHash"] = currentHash
        self.save_schedules()

        return "synced"


class SCR_CloudReviseAddSyncDialog(Window):
    """Modal dialog for picking a local source file and an existing Civillo target file/layer to pair up as a sync entry."""

    def __init__(self, macroFileFolder, civilloClient, orgNickname, projectId, existingEntry=None):

        with StreamReader(macroFileFolder + r"\SCR_CloudRevise_AddSync.xaml") as s:
            wpf.LoadComponent(self, s)

        ElementHost.EnableModelessKeyboardInterop(self)

        self.civilloClient = civilloClient
        self.orgNickname = orgNickname
        self.projectId = projectId
        self.existingEntry = existingEntry
        self.result = None

        existingLocalPath1 = existingEntry.get("localPath1", "") if existingEntry is not None else ""
        existingLocalPath2 = existingEntry.get("localPath2", "") if existingEntry is not None else ""

        self.localFolder1 = os.path.dirname(existingLocalPath1) if existingLocalPath1 else None
        self.localFolder2 = os.path.dirname(existingLocalPath2) if existingLocalPath2 else None

        # separate from localFolder1/2 (which drive the file list contents) - only used to seed the
        # folder browse dialogs, so a cleared file still leaves an empty list but Browse opens where you left off
        self.lastBrowseFolder1 = OptionsManager.GetString("SCR_CloudRevise_AddSync.lastlocalfolder1", "") or None
        self.lastBrowseFolder2 = OptionsManager.GetString("SCR_CloudRevise_AddSync.lastlocalfolder2", "") or None

        if existingEntry is not None:
            self.Title = "Edit Sync"
            self.okBtn.Content = "Apply"
        else:
            self.Title = "Add Sync"

        self.browseLocalFolderBtn1.Click += self.browse_local_folder1_clicked
        self.browseLocalFolderBtn2.Click += self.browse_local_folder2_clicked
        self.refreshCivilloFilesBtn.Click += self.refresh_civillo_files_clicked
        self.clearLocalFile1Btn.Click += self.clear_local_file1_clicked
        self.clearLocalFile2Btn.Click += self.clear_local_file2_clicked
        self.clearCivilloFileBtn.Click += self.clear_civillo_file_clicked
        self.okBtn.Click += self.ok_clicked
        self.cancelBtn.Click += self.cancel_clicked

        self.Loaded += self.restore_window_state
        self.Closing += self.save_window_state

        if self.orgNickname and self.projectId is not None:
            self.refresh_civillo_files_clicked(None, None)
        else:
            self.error.Content = "Select an organization and project in Settings first."

        if self.localFolder1:
            self.populate_local_files(self.localFileList1, self.localFolder1)
            self.select_item_by_text(self.localFileList1, os.path.basename(existingLocalPath1))

        if self.localFolder2:
            self.populate_local_files(self.localFileList2, self.localFolder2)
            self.select_item_by_text(self.localFileList2, os.path.basename(existingLocalPath2))


    # ---------- window position/size persistence ----------
    # SizeToContent="Width" (set in the xaml) auto-sizes the window on first run; once the user
    # has resized it, the saved size/position takes over and SizeToContent is switched off so it sticks.

    def restore_window_state(self, sender, e):
        savedWidth = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowwidth", 0)
        savedHeight = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowheight", 0)
        savedLeft = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowleft", -1)
        savedTop = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.windowtop", -1)

        if savedWidth > 0:
            self.SizeToContent = SizeToContent.Manual
            self.Width = savedWidth
        if savedHeight > 0:
            self.Height = savedHeight

        if savedLeft >= 0 and savedTop >= 0:
            self.Left = savedLeft
            self.Top = savedTop
            if not SCROptions._IsWindowOnAnyScreen(self):
                self.Left = 100
                self.Top = 100

        savedLeftStar = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.leftcolumnstar", 0)
        savedRightStar = OptionsManager.GetDouble("SCR_CloudRevise_AddSync.rightcolumnstar", 0)
        if savedLeftStar > 0 and savedRightStar > 0:
            self.rootGrid.ColumnDefinitions[0].Width = GridLength(savedLeftStar, GridUnitType.Star)
            self.rootGrid.ColumnDefinitions[2].Width = GridLength(savedRightStar, GridUnitType.Star)

    def save_window_state(self, sender, e):
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowwidth", self.Width)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowheight", self.Height)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowleft", self.Left)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.windowtop", self.Top)

        OptionsManager.SetValue("SCR_CloudRevise_AddSync.leftcolumnstar", self.rootGrid.ColumnDefinitions[0].Width.Value)
        OptionsManager.SetValue("SCR_CloudRevise_AddSync.rightcolumnstar", self.rootGrid.ColumnDefinitions[2].Width.Value)
        OptionsManager.SaveOptions()  # flush to disk now rather than relying on TBC's own exit-time flush

    def browse_local_folder1_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        startFolder1 = self.localFolder1 or self.lastBrowseFolder1
        if startFolder1 and os.path.isdir(startFolder1):
            dlg.SelectedPath = startFolder1
        if dlg.ShowDialog() == DialogResult.OK:
            self.localFolder1 = dlg.SelectedPath
            self.error.Content = ""
            self.populate_local_files(self.localFileList1, self.localFolder1)
            OptionsManager.SetValue("SCR_CloudRevise_AddSync.lastlocalfolder1", self.localFolder1)

    def browse_local_folder2_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        startFolder2 = self.localFolder2 or self.lastBrowseFolder2
        if startFolder2 and os.path.isdir(startFolder2):
            dlg.SelectedPath = startFolder2
        if dlg.ShowDialog() == DialogResult.OK:
            self.localFolder2 = dlg.SelectedPath
            self.error.Content = ""
            self.populate_local_files(self.localFileList2, self.localFolder2)
            OptionsManager.SetValue("SCR_CloudRevise_AddSync.lastlocalfolder2", self.localFolder2)

    def populate_local_files(self, listBox, folder):
        listBox.Items.Clear()
        if not folder:
            return
        try:
            for name in os.listdir(folder):
                full = os.path.join(folder, name)
                if os.path.isfile(full):
                    listBox.Items.Add(name)
        except Exception as ex:
            self.error.Content = str(ex)

    def select_item_by_text(self, listBox, targetName):
        if not targetName:
            return
        for item in listBox.Items:
            if str(item) == targetName:
                listBox.SelectedItem = item
                listBox.UpdateLayout()  # force layout so ScrollIntoView works before the window is shown
                listBox.ScrollIntoView(item)
                break

    def clear_local_file1_clicked(self, sender, e):
        self.localFileList1.SelectedItem = None

    def clear_local_file2_clicked(self, sender, e):
        self.localFileList2.SelectedItem = None

    def clear_civillo_file_clicked(self, sender, e):
        self.civilloFileList.SelectedItem = None

    def refresh_civillo_files_clicked(self, sender, e):
        self.civilloFileList.Items.Clear()
        self.error.Content = ""

        if not self.orgNickname or self.projectId is None:
            self.error.Content = "Select an organization and project in Settings first."
            return

        try:
            tree = self.civilloClient.get("/" + self.orgNickname + "/projects/" + str(self.projectId) + "/layer-directory")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        dump_civillo_layer_directory(tree, os.path.dirname(self.get_config_path()))

        files = []
        collect_civillo_files(tree, files)

        for f in files:
            folderText = f["folderPath"] if f["folderPath"].endswith("/") else f["folderPath"] + "/"

            folderRun = Run(folderText)
            folderRun.Foreground = SolidColorBrush(Colors.Black)  # matches the "->" separator color on the Schedules tab

            nameRun = Run(f["layerName"])
            nameRun.Foreground = SolidColorBrush(Colors.SteelBlue)  # matches the Civillo path color on the Schedules tab

            textBlock = TextBlock()
            textBlock.Inlines.Add(folderRun)
            textBlock.Inlines.Add(nameRun)

            item = ListBoxItem()
            item.Content = textBlock
            item.Tag = {"layerId": f["layerId"], "display": f["display"]}
            self.civilloFileList.Items.Add(item)

        # layers that exist in the project but were never organized into any /layer-directory folder -
        # by far the most common real case being a terrain/DTM/surface layer, which never appears in
        # /layer-directory at all regardless of how it was created (confirmed by hand against a real
        # project) - grouped under their own header rather than mixed into the folder tree above, since
        # they have no folder path to show
        try:
            unassigned = get_civillo_unassigned_layers(self.civilloClient, self.orgNickname, self.projectId, files)
        except Exception:
            unassigned = []  # non-fatal - the folder-tree list above already loaded fine either way

        if unassigned:
            unassigned.sort(key=lambda l: (l.get("name") or "").lower())

            titleRun = Run("Unassigned Layers")
            titleRun.FontWeight = FontWeights.Bold
            titleRun.FontStyle = FontStyles.Italic

            noteText = TextBlock()
            noteText.Text = "not organized into any folder - typically terrain/DTM/surface layers"
            noteText.FontSize = 10
            noteText.FontStyle = FontStyles.Italic
            noteText.Foreground = SolidColorBrush(Colors.Gray)
            noteText.TextWrapping = TextWrapping.Wrap

            headerText = TextBlock()
            headerText.Inlines.Add(titleRun)
            headerPanel = StackPanel()
            headerPanel.Children.Add(headerText)
            headerPanel.Children.Add(noteText)

            header = ListBoxItem()
            header.Content = headerPanel
            header.IsEnabled = False
            header.IsHitTestVisible = False
            header.Tag = None
            self.civilloFileList.Items.Add(header)

            for l in unassigned:
                nameRun = Run(l.get("name") or "(unnamed)")
                nameRun.Foreground = SolidColorBrush(Colors.SteelBlue)

                textBlock = TextBlock()
                textBlock.Inlines.Add(nameRun)

                item = ListBoxItem()
                item.Content = textBlock
                item.Tag = {"layerId": l.get("layerId"), "display": l.get("name") or "(unnamed)"}
                self.civilloFileList.Items.Add(item)

        if self.existingEntry is not None:
            targetLayerId = self.existingEntry.get("civilloLayerId")
            for item in self.civilloFileList.Items:
                if item.Tag is not None and item.Tag["layerId"] == targetLayerId:
                    self.civilloFileList.SelectedItem = item
                    self.civilloFileList.UpdateLayout()  # force layout so ScrollIntoView works before the window is shown
                    self.civilloFileList.ScrollIntoView(item)
                    break

    def ok_clicked(self, sender, e):
        local1Item = self.localFileList1.SelectedItem
        local2Item = self.localFileList2.SelectedItem
        civilloItem = self.civilloFileList.SelectedItem

        if local1Item is None or civilloItem is None:
            self.error.Content = "Select at least local file 1 and a Civillo target file."
            return

        localPath1 = os.path.join(self.localFolder1, str(local1Item))
        localPath2 = os.path.join(self.localFolder2, str(local2Item)) if local2Item is not None else ""

        self.result = {
            "localPath1": localPath1,
            "localPath2": localPath2,
            "civilloPath": civilloItem.Tag["display"],
            "civilloLayerId": civilloItem.Tag["layerId"],
            "orgNickname": self.orgNickname,
            "projectId": self.projectId,
        }
        self.Close()

    def cancel_clicked(self, sender, e):
        self.result = None
        self.Close()


class SCR_CloudReviseAddPropellerSyncDialog(Window):
    """Modal dialog for pairing a Propeller survey file with a Civillo target layer as a sync entry.
    The Trimble Connect zone is browse-only for now (region -> project -> folder tree) - it isn't wired
    into the Propeller/Civillo sync pairing yet, since SaveFileRemotely-based upload isn't built.

    The Propeller org/site/survey pickers live on the main window, not here - propellerContext carries
    the currently-selected survey's id/name info and its file list in from there (see
    SCR_CloudReviseDialog._current_propeller_context), since every deliverable file uses the same set of
    names for every survey and there's rarely a reason to change survey between adding sync entries."""

    def __init__(self, macroFileFolder, civilloClient, propellerContext, existingEntry=None):

        with StreamReader(macroFileFolder + r"\SCR_CloudRevise_AddPropellerSync.xaml") as s:
            wpf.LoadComponent(self, s)

        ElementHost.EnableModelessKeyboardInterop(self)

        self.civilloClient = civilloClient
        # {"orgId", "orgName", "siteId", "siteName", "surveyId", "surveyName", "files"} - the org/site/
        # survey pickers live on the main window now (org/site chosen once, applies to every sync entry
        # you add for that survey), so this dialog just shows the files for whatever survey was selected
        # there rather than fetching them itself
        self.propellerContext = propellerContext
        self.trimbleConnectClient = TrimbleConnectClient()
        self.existingEntry = existingEntry
        self.result = None

        self.tcProjectRoot = None
        self.tcFolderStack = []

        if existingEntry is not None:
            self.Title = "Edit Propeller Sync"
            self.okBtn.Content = "Apply"

        self.propContextLabel.Text = propellerContext["orgName"] + " / " + propellerContext["siteName"] + " / " + propellerContext["surveyName"]
        self.propFileList.SelectionChanged += self.prop_file_selection_changed

        self.civEnabledCheckbox.Checked += self.civ_enabled_changed
        self.civEnabledCheckbox.Unchecked += self.civ_enabled_changed
        self.civReloadBtn.Click += self.reload_civillo_orgs_clicked
        self.civOrgCombo.SelectionChanged += self.civillo_org_selection_changed
        self.civProjectCombo.SelectionChanged += self.civillo_project_selection_changed
        self.civMetersPerPixelBox.TextChanged += self.civ_ortho_pixel_size_changed
        self.civFormatCombo.SelectionChanged += self.civ_ortho_format_changed
        self.civMaxSizeGbBox.TextChanged += self.civ_max_size_changed
        self.civLazMethodVoxel.Checked += self.civ_laz_method_changed
        self.civLazMethodOctree.Checked += self.civ_laz_method_changed
        self.civLazMethodGrid.Checked += self.civ_laz_method_changed
        self.civLazMethodPdal.Checked += self.civ_laz_method_changed
        self.civLazSpacingBox.TextChanged += self.civ_laz_spacing_changed
        self.civCreateNewLayerCheckbox.Checked += self.civ_create_new_layer_changed
        self.civCreateNewLayerCheckbox.Unchecked += self.civ_create_new_layer_changed
        self.civNewLayerTitleBox.TextChanged += self.civ_new_layer_title_changed

        self.tcEnabledCheckbox.Checked += self.tc_enabled_changed
        self.tcEnabledCheckbox.Unchecked += self.tc_enabled_changed
        self.tcReloadBtn.Click += self.reload_tc_regions_clicked
        self.tcRegionCombo.SelectionChanged += self.tc_region_selection_changed
        self.tcProjectCombo.SelectionChanged += self.tc_project_selection_changed
        self.tcUpBtn.Click += self.tc_up_clicked
        self.tcFileList.MouseDoubleClick += self.tc_file_list_double_click
        self.tcMetersPerPixelBox.TextChanged += self.tc_ortho_pixel_size_changed
        self.tcFormatCombo.SelectionChanged += self.tc_ortho_format_changed
        self.tcMaxSizeGbBox.TextChanged += self.tc_max_size_changed
        self.tcLazMethodVoxel.Checked += self.tc_laz_method_changed
        self.tcLazMethodOctree.Checked += self.tc_laz_method_changed
        self.tcLazMethodGrid.Checked += self.tc_laz_method_changed
        self.tcLazMethodPdal.Checked += self.tc_laz_method_changed
        self.tcLazSpacingBox.TextChanged += self.tc_laz_spacing_changed

        self.okBtn.Click += self.ok_clicked
        self.cancelBtn.Click += self.cancel_clicked

        self.Loaded += self.restore_window_state
        self.Closing += self.save_window_state

        self._restore_ortho_settings(self.civMetersPerPixelBox, self.civFormatCombo, "civpixelsizemeters", "civoutputformat", "civilloPixelSizeMeters", "civilloOutputFormat")
        self._restore_ortho_settings(self.tcMetersPerPixelBox, self.tcFormatCombo, "tcpixelsizemeters", "tcoutputformat", "trimbleConnectPixelSizeMeters", "trimbleConnectOutputFormat")
        self._restore_max_size_setting(self.civMaxSizeGbBox, "civmaxsizegb", "civilloMaxSizeGb")
        self._restore_max_size_setting(self.tcMaxSizeGbBox, "tcmaxsizegb", "trimbleConnectMaxSizeGb")

        # PDAL only ever shows up as a choice when the main window's QGIS-PDAL Cloud resampling is both
        # enabled and actually found pdal.exe - otherwise the radio button stays Collapsed and out of
        # methodRadios entirely, so a previously-saved "pdal" preference falls back to voxel below rather
        # than restoring a hidden, uncheckable option (_restore_laz_settings' own methodRadios.get(...,
        # methodRadios["voxel"]) fallback handles that automatically once "pdal" is simply absent here)
        pdalAvailable = bool(self.propellerContext.get("pdalAvailable"))
        self.civLazMethodPdal.Visibility = Visibility.Visible if pdalAvailable else Visibility.Collapsed
        self.tcLazMethodPdal.Visibility = Visibility.Visible if pdalAvailable else Visibility.Collapsed

        self.civLazMethodRadios = {"voxel": self.civLazMethodVoxel, "octree": self.civLazMethodOctree, "grid": self.civLazMethodGrid}
        self.tcLazMethodRadios = {"voxel": self.tcLazMethodVoxel, "octree": self.tcLazMethodOctree, "grid": self.tcLazMethodGrid}
        if pdalAvailable:
            self.civLazMethodRadios["pdal"] = self.civLazMethodPdal
            self.tcLazMethodRadios["pdal"] = self.tcLazMethodPdal
        self._restore_laz_settings(self.civLazMethodRadios, self.civLazSpacingBox, "civlazmethod", "civlazspacing", "civilloLazResampleMethod", "civilloLazResampleSpacing")
        self._restore_laz_settings(self.tcLazMethodRadios, self.tcLazSpacingBox, "tclazmethod", "tclazspacing", "trimbleConnectLazResampleMethod", "trimbleConnectLazResampleSpacing")

        # "create additional new layer" defaults to off (unlike civ/tcEnabledCheckbox, which default on) -
        # this is an opt-in extra, not a normal target, so a schedule entry that's never touched it
        # shouldn't suddenly start creating layers
        entryCreateNew = self.existingEntry.get("civilloCreateNewLayer") if self.existingEntry is not None else None
        if entryCreateNew is not None:
            self.civCreateNewLayerCheckbox.IsChecked = bool(entryCreateNew)
        else:
            savedCreateNew = self.get_saved("civcreatenewlayer")
            self.civCreateNewLayerCheckbox.IsChecked = bool(savedCreateNew) and savedCreateNew != "False"

        entryNewLayerTitle = self.existingEntry.get("civilloNewLayerTitleTemplate") if self.existingEntry is not None else None
        self.civNewLayerTitleBox.Text = entryNewLayerTitle if entryNewLayerTitle else (self.get_saved("civnewlayertitle") or "{YYMMDD} ")

        self.civEnabledCheckbox.IsChecked = self._resolve_enabled_checkbox("civilloEnabled", "civenabled")
        self.tcEnabledCheckbox.IsChecked = self._resolve_enabled_checkbox("trimbleConnectEnabled", "tcenabled")

        self.populate_propeller_file_list()
        self.reload_civillo_orgs_clicked(None, None)
        self.reload_tc_regions_clicked(None, None)


    # ---------- window position/size persistence ----------

    def restore_window_state(self, sender, e):
        savedWidth = OptionsManager.GetDouble("SCR_CloudRevise_AddPropellerSync.windowwidth", 0)
        savedHeight = OptionsManager.GetDouble("SCR_CloudRevise_AddPropellerSync.windowheight", 0)
        savedLeft = OptionsManager.GetDouble("SCR_CloudRevise_AddPropellerSync.windowleft", -1)
        savedTop = OptionsManager.GetDouble("SCR_CloudRevise_AddPropellerSync.windowtop", -1)

        if savedWidth > 0:
            self.Width = savedWidth
        if savedHeight > 0:
            self.Height = savedHeight

        if savedLeft >= 0 and savedTop >= 0:
            self.Left = savedLeft
            self.Top = savedTop
            if not SCROptions._IsWindowOnAnyScreen(self):
                self.Left = 100
                self.Top = 100

        col0 = OptionsManager.GetDouble("SCR_CloudRevise_AddPropellerSync.col0star", 0)
        col2 = OptionsManager.GetDouble("SCR_CloudRevise_AddPropellerSync.col2star", 0)
        col4 = OptionsManager.GetDouble("SCR_CloudRevise_AddPropellerSync.col4star", 0)
        if col0 > 0 and col2 > 0 and col4 > 0:
            self.rootGrid.ColumnDefinitions[0].Width = GridLength(col0, GridUnitType.Star)
            self.rootGrid.ColumnDefinitions[2].Width = GridLength(col2, GridUnitType.Star)
            self.rootGrid.ColumnDefinitions[4].Width = GridLength(col4, GridUnitType.Star)

    def save_window_state(self, sender, e):
        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync.windowwidth", self.Width)
        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync.windowheight", self.Height)
        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync.windowleft", self.Left)
        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync.windowtop", self.Top)

        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync.col0star", self.rootGrid.ColumnDefinitions[0].Width.Value)
        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync.col2star", self.rootGrid.ColumnDefinitions[2].Width.Value)
        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync.col4star", self.rootGrid.ColumnDefinitions[4].Width.Value)
        OptionsManager.SaveOptions()  # flush to disk now rather than relying on TBC's own exit-time flush


    # ---------- remembered combo selections ----------

    def get_saved(self, key):
        return OptionsManager.GetString("SCR_CloudRevise_AddPropellerSync." + key, "")

    def save_selected(self, key, value):
        OptionsManager.SetValue("SCR_CloudRevise_AddPropellerSync." + key, str(value) if value is not None else "")

    def select_preferred(self, combo, entryKey, optionKey):
        # picks the item whose Tag matches (in order of preference) the entry being edited, then the
        # last value remembered from a previous use of this dialog, then just falls back to index 0
        preferred = self.existingEntry.get(entryKey) if self.existingEntry is not None else None
        if preferred is None:
            preferred = self.get_saved(optionKey) or None

        target = None
        if preferred is not None:
            for item in combo.Items:
                if str(item.Tag) == str(preferred):
                    target = item
                    break

        if target is not None:
            combo.SelectedItem = target
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0

    def select_preferred_by_content(self, combo, entryKey, optionKey):
        # same idea as select_preferred, but matches on the item's displayed text rather than its Tag -
        # Trimble Connect regions/projects have no stable id we can persist (region.ID is always None),
        # so their display name is the only usable key
        preferred = self.existingEntry.get(entryKey) if self.existingEntry is not None else None
        if preferred is None:
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


    # ---------- Propeller side ----------

    def populate_propeller_file_list(self):
        self.propFileList.Items.Clear()
        self.error.Content = ""

        files = self.propellerContext.get("files") or []

        def normalize_crs_label(text):
            # Propeller sometimes puts a raw code straight in the name's parens too (seen: "EPSG::4326",
            # double colon) - fold that to the same "EPSG:nnnn" shape everything else here uses, so it
            # doesn't end up as its own group purely over a formatting difference
            codeMatch = re.match(r"^EPSG:{1,2}(\d+)$", text.strip(), re.IGNORECASE)
            if not codeMatch:
                return text
            code = codeMatch.group(1)
            return "EPSG:" + code + " (WGS84)" if code == "4326" else "EPSG:" + code

        # scan once for a friendly (non-code-shaped) CRS name already present among files that aren't
        # WGS84-flagged - if it's there, it's Propeller's own name for the site's default CRS, so files
        # that fall back to the site's own EPSG code below (LAZ/point clouds notably, which don't get a
        # "(...)" annotation in "name" at all) can share that same friendly group instead of splitting
        # into a separate "EPSG:nnnn" group purely because their file happened to lack a name annotation
        friendlySiteCrsName = None
        for f in files:
            if f.get("is_wgs84"):
                continue
            nameMatch = re.search(r"\(([^)]*)\)", f.get("name", ""))
            if nameMatch and not re.match(r"^EPSG:{1,2}\d+$", nameMatch.group(1).strip(), re.IGNORECASE):
                friendlySiteCrsName = nameMatch.group(1)
                break

        def crs_group_label(f):
            # orthophotos carry a human-readable CRS name right in the API's own "name" field, e.g.
            # "Ortho TIFF (GDA2020 / MGA zone 56)" - prefer that when present, it reads better than a
            # bare code. Everything else (LAZ point clouds notably) doesn't get that annotation, so falls
            # back to a friendly name spotted elsewhere in this same list, or the site's own
            # coordinate_reference_system_pointer as a plain "EPSG:nnnn" label if there isn't one - unlike
            # guessing a name out of the filename (which only worked for Australia's "zone_NN"
            # convention), a code straight from the site record is correct everywhere in the world.
            # is_wgs84 overrides both: that flag means this particular file's own coordinates are WGS84
            # regardless of what CRS the site itself is configured with.
            match = re.search(r"\(([^)]*)\)", f.get("name", ""))
            if match:
                return normalize_crs_label(match.group(1))
            if f.get("is_wgs84"):
                return "EPSG:4326 (WGS84)"
            return friendlySiteCrsName or self.propellerContext.get("siteCrsLabel") or "Unknown Coordinate System"

        sortedFiles = sorted(files, key=lambda f: (crs_group_label(f), f.get("name", ""), f.get("type", "")))

        lastGroup = None
        for f in sortedFiles:
            group = crs_group_label(f)
            if group != lastGroup:
                header = ListBoxItem()
                header.Content = group
                header.FontWeight = FontWeights.Bold
                header.IsEnabled = False
                header.IsHitTestVisible = False
                header.Tag = None
                self.propFileList.Items.Add(header)
                lastGroup = group

            # orthophoto (GDAL resample) and LAZ (point cloud -> LandXML surface) are the only things the
            # run flow actually supports right now - color everything else red so it's obvious at a
            # glance before picking it
            isSupported = f.get("type") == "orthophoto" or f.get("format") == "laz"

            nameRun = Run(f.get("name", "(unnamed)") + " (" + f.get("type", "?") + ")")
            nameRun.Foreground = SolidColorBrush(Colors.DarkGreen if isSupported else Colors.Red)

            textBlock = TextBlock()
            textBlock.Inlines.Add(nameRun)

            item = ListBoxItem()
            item.Content = textBlock
            item.Tag = f
            self.propFileList.Items.Add(item)

        if self.existingEntry is not None:
            # matched by name, not by url/id - an entry no longer remembers a specific file's url (files
            # aren't tied to one survey any more), just its name string, which is what Propeller keeps
            # consistent across every survey
            targetName = self.existingEntry.get("propellerFileRawName")
            for item in self.propFileList.Items:
                if item.Tag is not None and item.Tag.get("name") == targetName:
                    self.propFileList.SelectedItem = item
                    self.propFileList.UpdateLayout()
                    self.propFileList.ScrollIntoView(item)
                    break

        self.prop_file_selection_changed(None, None)

    def prop_file_selection_changed(self, sender, e):
        # the pixel/m + format resample options only make sense for an orthophoto - other deliverable
        # types (terrain_dtm/terrain_dsm/pointcloud) aren't images GDAL can resample this way, and the
        # LAZ resample options only make sense for a LAZ point cloud. Each platform's panels also stay
        # off while that platform's own checkbox is unchecked.
        self._update_ortho_panels_enabled()

    def _update_ortho_panels_enabled(self):
        item = self.propFileList.SelectedItem
        isOrtho = item is not None and item.Tag is not None and item.Tag.get("type") == "orthophoto"
        isLaz = item is not None and item.Tag is not None and item.Tag.get("format") == "laz"
        civOrthoEnabled = isOrtho and bool(self.civEnabledCheckbox.IsChecked)
        tcOrthoEnabled = isOrtho and bool(self.tcEnabledCheckbox.IsChecked)
        civLazEnabled = isLaz and bool(self.civEnabledCheckbox.IsChecked)
        tcLazEnabled = isLaz and bool(self.tcEnabledCheckbox.IsChecked)
        # each pair set from the same locally-computed bool, rather than one read back from the other's
        # (possibly-coerced-by-the-ancestor-chain) IsEnabled getter, so both panels land in the exact same
        # state regardless of WPF's IsEnabled coercion/binding-update timing
        self.civOrthoOptionsPanel.IsEnabled = civOrthoEnabled
        self.tcOrthoOptionsPanel.IsEnabled = tcOrthoEnabled
        self.civLazOptionsPanel.IsEnabled = civLazEnabled
        self.tcLazOptionsPanel.IsEnabled = tcLazEnabled
        # same "max GB" box/setting serves both file types - which one it limits just follows whichever
        # type is actually selected, so it doesn't need its own separate LAZ copy
        self.civMaxSizePanel.IsEnabled = civOrthoEnabled or civLazEnabled
        self.tcMaxSizePanel.IsEnabled = tcOrthoEnabled or tcLazEnabled

        # hidden (not just disabled) when they don't apply to the selected file's type at all - a LAZ
        # file can never use pixel-size/format settings and vice versa, so leaving them visible-but-greyed
        # just clutters the dialog with controls that can never do anything for this selection. Based on
        # file type only (not the per-platform enabled checkboxes, unlike IsEnabled above) - turning
        # Civillo off shouldn't hide settings that are still relevant for Trimble Connect, or vice versa.
        orthoVisibility = Visibility.Visible if isOrtho else Visibility.Collapsed
        lazVisibility = Visibility.Visible if isLaz else Visibility.Collapsed
        maxSizeVisibility = Visibility.Visible if (isOrtho or isLaz) else Visibility.Collapsed
        self.civOrthoOptionsPanel.Visibility = orthoVisibility
        self.tcOrthoOptionsPanel.Visibility = orthoVisibility
        self.civLazOptionsPanel.Visibility = lazVisibility
        self.tcLazOptionsPanel.Visibility = lazVisibility
        self.civMaxSizePanel.Visibility = maxSizeVisibility
        self.tcMaxSizePanel.Visibility = maxSizeVisibility
        # "create additional new layer" only makes sense for LAZ->terrain layers (the whole feature exists
        # because terrain layers can't be discovered/revised via /layer-directory) - a TIFF/JPEG orthophoto
        # always gets its own revisable layer, so the checkbox is meaningless for it
        self.civCreateNewLayerPanel.Visibility = lazVisibility

    def _restore_ortho_settings(self, box, combo, pixelKey, formatKey, entryPixelKey, entryFormatKey):
        # existingEntry (when editing) wins over the globally-remembered last-used value, same precedence
        # select_preferred() already uses elsewhere in this dialog - otherwise editing an entry shows
        # whatever was last typed into ANY entry's box rather than this entry's own saved setting, which
        # looks indistinguishable from "the setting didn't save" even though it actually did. Editing an
        # entry that deliberately never had a pixel size (blank = "don't resample unless max-GB forces
        # it" - see _prepare_platform_file) must show blank too, not fall through to some other entry's
        # remembered value - that's the "Add" dialog's job only, when there's no existingEntry yet at all.
        if self.existingEntry is not None:
            entryPixelSize = self.existingEntry.get(entryPixelKey)
            # "%.10g" rather than str()/repr() - see the comment on the same pattern in _restore_max_size_setting
            box.Text = ("%.10g" % entryPixelSize) if entryPixelSize is not None else ""
        else:
            savedPixelSize = self.get_saved(pixelKey)
            if savedPixelSize:
                box.Text = savedPixelSize

        entryFormat = self.existingEntry.get(entryFormatKey) if self.existingEntry is not None else None
        savedFormat = entryFormat if entryFormat else self.get_saved(formatKey)
        if savedFormat:
            for item in combo.Items:
                if str(item.Content) == savedFormat:
                    combo.SelectedItem = item
                    break

    def _restore_max_size_setting(self, box, optionKey, entryKey):
        # same existingEntry-wins-over-remembered precedence as _restore_ortho_settings, and the same
        # fix: editing an entry that never had a max-GB limit must show blank, not some other entry's
        # last-remembered value - the global fallback only applies to the "Add" dialog (no existingEntry)
        if self.existingEntry is not None:
            entryMaxSize = self.existingEntry.get(entryKey)
            # "%.10g" rather than str()/repr() - IronPython's float-to-str doesn't always give the
            # shortest round-trip representation CPython's does, so a value entered as "0.03" can come
            # back out of storage as the same double but stringify as "0.029999999999999999"; 10
            # significant digits is comfortably past any user-meaningful precision for a GB size while
            # trimming that noise off
            box.Text = ("%.10g" % entryMaxSize) if entryMaxSize is not None else ""
        else:
            box.Text = self.get_saved(optionKey) or ""

    def _restore_laz_settings(self, methodRadios, spacingBox, methodOptionKey, spacingOptionKey, entryMethodKey, entrySpacingKey):
        # methodRadios: {"voxel": radioButton, "octree": radioButton, "grid": radioButton} - same
        # existingEntry-wins-over-last-remembered precedence as _restore_ortho_settings/_restore_max_size_setting
        entryMethod = self.existingEntry.get(entryMethodKey) if self.existingEntry is not None else None
        savedMethod = entryMethod if entryMethod else (self.get_saved(methodOptionKey) or "voxel")
        radio = methodRadios.get(savedMethod, methodRadios["voxel"])
        radio.IsChecked = True

        entrySpacing = self.existingEntry.get(entrySpacingKey) if self.existingEntry is not None else None
        # "%.10g" rather than str()/repr() - see the comment on the same pattern in _restore_max_size_setting
        savedSpacing = ("%.10g" % entrySpacing) if entrySpacing is not None else (self.get_saved(spacingOptionKey) or "10")
        spacingBox.Text = savedSpacing

    def civ_ortho_pixel_size_changed(self, sender, e):
        self.save_selected("civpixelsizemeters", self.civMetersPerPixelBox.Text)

    def civ_ortho_format_changed(self, sender, e):
        item = self.civFormatCombo.SelectedItem
        if item is not None:
            self.save_selected("civoutputformat", str(item.Content))

    def tc_ortho_pixel_size_changed(self, sender, e):
        self.save_selected("tcpixelsizemeters", self.tcMetersPerPixelBox.Text)

    def tc_ortho_format_changed(self, sender, e):
        item = self.tcFormatCombo.SelectedItem
        if item is not None:
            self.save_selected("tcoutputformat", str(item.Content))

    def civ_max_size_changed(self, sender, e):
        self.save_selected("civmaxsizegb", self.civMaxSizeGbBox.Text)

    def tc_max_size_changed(self, sender, e):
        self.save_selected("tcmaxsizegb", self.tcMaxSizeGbBox.Text)

    def civ_laz_method_changed(self, sender, e):
        if sender.IsChecked:
            method = next(k for k, rb in self.civLazMethodRadios.items() if rb is sender)
            self.save_selected("civlazmethod", method)

    def civ_laz_spacing_changed(self, sender, e):
        self.save_selected("civlazspacing", self.civLazSpacingBox.Text)

    def civ_create_new_layer_changed(self, sender, e):
        self.save_selected("civcreatenewlayer", str(bool(self.civCreateNewLayerCheckbox.IsChecked)))

    def civ_new_layer_title_changed(self, sender, e):
        self.save_selected("civnewlayertitle", self.civNewLayerTitleBox.Text)

    def tc_laz_method_changed(self, sender, e):
        if sender.IsChecked:
            method = next(k for k, rb in self.tcLazMethodRadios.items() if rb is sender)
            self.save_selected("tclazmethod", method)

    def tc_laz_spacing_changed(self, sender, e):
        self.save_selected("tclazspacing", self.tcLazSpacingBox.Text)

    def _resolve_enabled_checkbox(self, entryKey, optionKey):
        # existingEntry (when editing) wins, then whatever was remembered from last time, then True
        if self.existingEntry is not None and entryKey in self.existingEntry:
            return bool(self.existingEntry[entryKey])
        saved = self.get_saved(optionKey)
        if saved:
            return saved != "False"
        return True

    def civ_enabled_changed(self, sender, e):
        self.save_selected("civenabled", str(bool(self.civEnabledCheckbox.IsChecked)))
        self._update_ortho_panels_enabled()

    def tc_enabled_changed(self, sender, e):
        self.save_selected("tcenabled", str(bool(self.tcEnabledCheckbox.IsChecked)))
        self._update_ortho_panels_enabled()

    # ---------- Civillo side ----------

    def reload_civillo_orgs_clicked(self, sender, e):
        self.error.Content = ""
        self.civOrgCombo.Items.Clear()
        self.civProjectCombo.Items.Clear()
        self.civFileList.Items.Clear()

        try:
            orgs = self.civilloClient.get("/applications")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for org in orgs:
            item = ComboBoxItem()
            item.Content = org["name"]
            item.Tag = org["nickname"]
            self.civOrgCombo.Items.Add(item)

        self.select_preferred(self.civOrgCombo, "civilloOrgNickname", "selectedcivorgnickname")

    def civillo_org_selection_changed(self, sender, e):
        self.civProjectCombo.Items.Clear()
        self.civFileList.Items.Clear()
        self.error.Content = ""

        orgItem = self.civOrgCombo.SelectedItem
        if orgItem is None:
            return
        self.save_selected("selectedcivorgnickname", orgItem.Tag)

        try:
            projects = self.civilloClient.get("/" + orgItem.Tag + "/projects")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for proj in projects:
            item = ComboBoxItem()
            item.Content = proj["name"]
            item.Tag = proj["id"]
            self.civProjectCombo.Items.Add(item)

        self.select_preferred(self.civProjectCombo, "civilloProjectId", "selectedcivprojectid")

    def civillo_project_selection_changed(self, sender, e):
        self.civFileList.Items.Clear()
        self.error.Content = ""

        orgItem = self.civOrgCombo.SelectedItem
        projectItem = self.civProjectCombo.SelectedItem
        if orgItem is None or projectItem is None:
            return
        self.save_selected("selectedcivprojectid", projectItem.Tag)

        try:
            tree = self.civilloClient.get("/" + orgItem.Tag + "/projects/" + str(projectItem.Tag) + "/layer-directory")
        except Exception as ex:
            self.error.Content = str(ex)
            return

        # this dialog has no get_config_path() of its own (that's on the main window) - same folder,
        # built the same way, since it's just for a diagnostic file dump
        dump_civillo_layer_directory(tree, os.path.join(os.environ.get("APPDATA"), "SCR Macros", "SCR_CloudRevise"))

        files = []
        collect_civillo_files(tree, files)

        for f in files:
            folderText = f["folderPath"] if f["folderPath"].endswith("/") else f["folderPath"] + "/"

            folderRun = Run(folderText)
            folderRun.Foreground = SolidColorBrush(Colors.Black)

            nameRun = Run(f["layerName"])
            nameRun.Foreground = SolidColorBrush(Colors.SteelBlue)

            textBlock = TextBlock()
            textBlock.Inlines.Add(folderRun)
            textBlock.Inlines.Add(nameRun)

            item = ListBoxItem()
            item.Content = textBlock
            item.Tag = {"layerId": f["layerId"], "display": f["display"]}
            self.civFileList.Items.Add(item)

        # see the matching block in refresh_civillo_files_clicked - same "not in any folder" elimination,
        # covers terrain/DTM/surface layers, which never show up in /layer-directory at all
        try:
            unassigned = get_civillo_unassigned_layers(self.civilloClient, orgItem.Tag, projectItem.Tag, files)
        except Exception:
            unassigned = []

        if unassigned:
            unassigned.sort(key=lambda l: (l.get("name") or "").lower())

            titleRun = Run("Unassigned Layers")
            titleRun.FontWeight = FontWeights.Bold
            titleRun.FontStyle = FontStyles.Italic

            noteText = TextBlock()
            noteText.Text = "not organized into any folder - typically terrain/DTM/surface layers"
            noteText.FontSize = 10
            noteText.FontStyle = FontStyles.Italic
            noteText.Foreground = SolidColorBrush(Colors.Gray)
            noteText.TextWrapping = TextWrapping.Wrap

            headerText = TextBlock()
            headerText.Inlines.Add(titleRun)
            headerPanel = StackPanel()
            headerPanel.Children.Add(headerText)
            headerPanel.Children.Add(noteText)

            header = ListBoxItem()
            header.Content = headerPanel
            header.IsEnabled = False
            header.IsHitTestVisible = False
            header.Tag = None
            self.civFileList.Items.Add(header)

            for l in unassigned:
                nameRun = Run(l.get("name") or "(unnamed)")
                nameRun.Foreground = SolidColorBrush(Colors.SteelBlue)

                textBlock = TextBlock()
                textBlock.Inlines.Add(nameRun)

                item = ListBoxItem()
                item.Content = textBlock
                item.Tag = {"layerId": l.get("layerId"), "display": l.get("name") or "(unnamed)"}
                self.civFileList.Items.Add(item)

        if self.existingEntry is not None:
            targetLayerId = self.existingEntry.get("civilloLayerId")
            for item in self.civFileList.Items:
                if item.Tag is not None and item.Tag["layerId"] == targetLayerId:
                    self.civFileList.SelectedItem = item
                    self.civFileList.UpdateLayout()
                    self.civFileList.ScrollIntoView(item)
                    break


    # ---------- Trimble Connect side ----------
    # browse-only: region -> project -> folder tree, navigated one level at a time via double-click/Up,
    # since GetFileList only returns one folder's immediate children per call (no single "whole tree" call
    # like Civillo's layer-directory - eagerly flattening a project with tens of thousands of files would
    # mean tens of thousands of round trips)

    def reload_tc_regions_clicked(self, sender, e):
        self.error.Content = ""
        self.tcRegionCombo.Items.Clear()
        self.tcProjectCombo.Items.Clear()
        self.tcFileList.Items.Clear()
        self.tcProjectRoot = None
        self.tcFolderStack = []

        try:
            regions = self.trimbleConnectClient.get_regions()
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for r in sorted(regions, key=lambda x: str(x.FileName)):
            item = ComboBoxItem()
            item.Content = str(r.FileName)
            item.Tag = r
            self.tcRegionCombo.Items.Add(item)

        self.select_preferred_by_content(self.tcRegionCombo, "trimbleConnectRegionName", "selectedtcregionname")

    def tc_region_selection_changed(self, sender, e):
        self.tcProjectCombo.Items.Clear()
        self.tcFileList.Items.Clear()
        self.tcProjectRoot = None
        self.tcFolderStack = []
        self.error.Content = ""

        regionItem = self.tcRegionCombo.SelectedItem
        if regionItem is None:
            return
        self.save_selected("selectedtcregionname", regionItem.Content)

        try:
            projects = self.trimbleConnectClient.get_children(regionItem.Tag)
        except Exception as ex:
            self.error.Content = str(ex)
            return

        for p in sorted(projects, key=lambda x: str(x.FileName)):
            item = ComboBoxItem()
            item.Content = str(p.FileName)
            item.Tag = p
            self.tcProjectCombo.Items.Add(item)

        self.select_preferred_by_content(self.tcProjectCombo, "trimbleConnectProjectName", "selectedtcprojectname")

    def tc_project_selection_changed(self, sender, e):
        self.tcFileList.Items.Clear()
        self.tcFolderStack = []
        self.error.Content = ""

        projectItem = self.tcProjectCombo.SelectedItem
        if projectItem is None:
            self.tcProjectRoot = None
            return
        self.save_selected("selectedtcprojectname", projectItem.Content)

        self.tcProjectRoot = projectItem.Tag
        self.restore_tc_target_selection()

    def restore_tc_target_selection(self):
        # walks self.existingEntry's saved folder path (a list of folder names from the project root down
        # to the folder containing the target) so Edit lands back where the target was picked from, then
        # selects the target item itself (which may be a folder or a file) in that folder's listing
        targetName = self.existingEntry.get("trimbleConnectTargetName") if self.existingEntry is not None else None
        if not targetName:
            self.refresh_tc_file_list()
            return

        folderPath = self.existingEntry.get("trimbleConnectFolderPath") or []
        self.tcFolderStack = []
        current = self.tcProjectRoot
        for folderName in folderPath:
            try:
                children = self.trimbleConnectClient.get_children(current)
            except Exception:
                break
            match = None
            for c in children:
                if str(c.FileName) == folderName:
                    match = c
                    break
            if match is None:
                break
            self.tcFolderStack.append(match)
            current = match

        self.refresh_tc_file_list()

        for item in self.tcFileList.Items:
            if str(item.Tag.FileName) == targetName:
                self.tcFileList.SelectedItem = item
                self.tcFileList.UpdateLayout()
                self.tcFileList.ScrollIntoView(item)
                break

    def tc_current_folder(self):
        return self.tcFolderStack[-1] if self.tcFolderStack else self.tcProjectRoot

    def resolve_tc_target(self):
        # (targetObject, isFolder, containingFolderPathNames) for whatever counts as the Trimble Connect
        # target right now: an explicit row picked in tcFileList if there is one; otherwise the folder
        # currently being browsed - so navigating into a folder and not picking a row inside it still
        # counts as "use this folder", instead of silently meaning nothing was selected
        tcTargetItem = self.tcFileList.SelectedItem
        if tcTargetItem is not None:
            return tcTargetItem.Tag, bool(tcTargetItem.Tag.IsFolder), [str(f.FileName) for f in self.tcFolderStack]

        if self.tcFolderStack:
            containingPath = [str(f.FileName) for f in self.tcFolderStack[:-1]]
            return self.tcFolderStack[-1], True, containingPath

        if self.tcProjectRoot is not None:
            return self.tcProjectRoot, True, []

        return None, None, None

    def refresh_tc_file_list(self):
        self.tcFileList.Items.Clear()
        self.tcPathLabel.Text = "/" + "/".join(str(f.FileName) for f in self.tcFolderStack)

        folder = self.tc_current_folder()
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
            self.tcFileList.Items.Add(item)

    def tc_file_list_double_click(self, sender, e):
        item = self.tcFileList.SelectedItem
        if item is None or not item.Tag.IsFolder:
            return
        self.tcFolderStack.append(item.Tag)
        self.refresh_tc_file_list()

    def tc_up_clicked(self, sender, e):
        if self.tcFolderStack:
            self.tcFolderStack.pop()
            self.refresh_tc_file_list()


    # ---------- ok/cancel ----------

    def ok_clicked(self, sender, e):
        propFileItem = self.propFileList.SelectedItem
        civOrgItem = self.civOrgCombo.SelectedItem
        civProjectItem = self.civProjectCombo.SelectedItem
        civFileItem = self.civFileList.SelectedItem if self.civEnabledCheckbox.IsChecked else None
        propIsLaz = propFileItem is not None and propFileItem.Tag is not None and propFileItem.Tag.get("format") == "laz"
        civCreateNew = bool(self.civCreateNewLayerCheckbox.IsChecked) and bool(self.civEnabledCheckbox.IsChecked) and propIsLaz
        tcRegionItem = self.tcRegionCombo.SelectedItem
        tcProjectItem = self.tcProjectCombo.SelectedItem
        tcTarget, tcIsFolder, tcFolderPath = self.resolve_tc_target() if self.tcEnabledCheckbox.IsChecked else (None, None, None)

        if propFileItem is None or (civFileItem is None and not civCreateNew and tcTarget is None):
            self.error.Content = "Select a Propeller file and a Civillo (existing layer or \"create additional new layer\") or Trimble Connect target (with its checkbox enabled)."
            return

        # propFileItem.Content is a colored TextBlock now (see propeller_survey_selection_changed), not
        # plain text - rebuild the display string from the underlying file dict instead of str()'ing it
        propFileDisplayText = propFileItem.Tag.get("name", "(unnamed)") + " (" + propFileItem.Tag.get("type", "?") + ")"

        # recomputed the same way _update_ortho_panels_enabled() derives it, rather than read back from
        # civOrthoOptionsPanel.IsEnabled/tcOrthoOptionsPanel.IsEnabled - those getters go through WPF's
        # IsEnabled coercion against the ancestor chain, which bit us once already (the max-size panel
        # showing disabled when editing); recomputing from the same source condition sidesteps that
        # class of bug entirely rather than trusting a coerced readback at save time too
        isOrtho = propFileItem.Tag is not None and propFileItem.Tag.get("type") == "orthophoto"
        isLaz = propFileItem.Tag is not None and propFileItem.Tag.get("format") == "laz"

        targetLabels = []
        result = {
            # site is the only propeller id kept - org and survey aren't stored at all: a schedule entry
            # isn't pinned to one survey any more, and org is only ever needed live (to browse to a site),
            # never at run time. propellerFileUrl isn't kept either - the run flow re-resolves the file by
            # its name string (propellerFileRawName) against whatever survey is currently selected, since
            # a url would tie the entry to the specific survey it was added under and Propeller's file
            # names are consistent across every survey anyway.
            "propellerSiteId": self.propellerContext["siteId"],
            "propellerSiteName": self.propellerContext["siteName"],
            "propellerFileName": propFileDisplayText,
            "propellerFileRawName": propFileItem.Tag.get("name"),
            "propellerFileType": propFileItem.Tag.get("type"),
            "propellerFileFormat": propFileItem.Tag.get("format"),
            "propellerFileSizeBytes": propFileItem.Tag.get("size_bytes"),
        }

        if civFileItem is not None or civCreateNew:
            result["civilloOrgNickname"] = civOrgItem.Tag
            result["civilloOrgName"] = str(civOrgItem.Content)
            result["civilloProjectId"] = civProjectItem.Tag
            result["civilloProjectName"] = str(civProjectItem.Content)

            if civFileItem is not None:
                result["civilloLayerId"] = civFileItem.Tag["layerId"]
                result["civilloDisplay"] = civFileItem.Tag["display"]
                targetLabels.append("Civillo: " + civFileItem.Tag["display"])

            result["civilloCreateNewLayer"] = civCreateNew
            if civCreateNew:
                result["civilloNewLayerTitleTemplate"] = self.civNewLayerTitleBox.Text
                targetLabels.append("Civillo: create new layer '" + self.civNewLayerTitleBox.Text + "'")

            if isOrtho:
                # format is independent of whether a pixel size was entered at all - a blank m/pixel box
                # still means "keep native resolution" for whichever format was picked (JPEG still needs
                # an actual re-encode, see _prepare_platform_file), so it must be saved even when the
                # pixel size itself is blank/invalid, not skipped along with it in the same try block
                civFormatItem = self.civFormatCombo.SelectedItem
                result["civilloOutputFormat"] = str(civFormatItem.Content) if civFormatItem is not None else "TIFF"
                try:
                    result["civilloPixelSizeMeters"] = float(self.civMetersPerPixelBox.Text)
                except Exception:
                    pass  # invalid/blank pixel size - leave unset, the run flow keeps native resolution

            if isLaz:
                result["civilloLazResampleMethod"] = next(k for k, rb in self.civLazMethodRadios.items() if rb.IsChecked)
                try:
                    result["civilloLazResampleSpacing"] = float(self.civLazSpacingBox.Text)
                except Exception:
                    pass  # invalid spacing - leave unset, the run flow falls back to its own default

            if (isOrtho or isLaz) and self.civMaxSizeGbBox.Text.strip():
                try:
                    result["civilloMaxSizeGb"] = float(self.civMaxSizeGbBox.Text)
                except Exception:
                    pass  # invalid max size - leave unset, the run flow just skips the size-limit step

        if tcTarget is not None:
            # a folder target has no file yet - the actual filename is derived from the source file at
            # sync time and placed inside this folder, rather than replacing a specific existing file.
            # trimbleConnectFolderPath (containing chain) is kept for restoring the picker's navigation on
            # Edit; trimbleConnectUploadFolderPath (containing chain + the target itself, if it's a folder)
            # is the actual path the run flow walks to reach the folder to upload into.
            result["trimbleConnectRegionName"] = str(tcRegionItem.Content) if tcRegionItem is not None else ""
            result["trimbleConnectProjectName"] = str(tcProjectItem.Content) if tcProjectItem is not None else ""
            result["trimbleConnectFolderPath"] = tcFolderPath
            result["trimbleConnectTargetName"] = str(tcTarget.FileName)
            result["trimbleConnectTargetIsFolder"] = bool(tcIsFolder)
            # when the target IS the project root itself (topmost folder, nothing navigated/selected),
            # tcTarget.FileName is the *project's own name*, not a subfolder to descend into - only
            # append it when the target is an actual subfolder distinct from the project root
            isProjectRootTarget = tcTarget is self.tcProjectRoot
            result["trimbleConnectUploadFolderPath"] = tcFolderPath + ([result["trimbleConnectTargetName"]] if (tcIsFolder and not isProjectRootTarget) else [])
            tcPath = "/".join(tcFolderPath + [result["trimbleConnectTargetName"]])
            targetLabels.append("Trimble Connect: " + tcPath + (" (folder)" if tcIsFolder else ""))

            if isOrtho:
                # format is independent of whether a pixel size was entered - see the matching Civillo
                # block above for why this must not be skipped along with a blank/invalid pixel size
                tcFormatItem = self.tcFormatCombo.SelectedItem
                result["trimbleConnectOutputFormat"] = str(tcFormatItem.Content) if tcFormatItem is not None else "TIFF"
                try:
                    result["trimbleConnectPixelSizeMeters"] = float(self.tcMetersPerPixelBox.Text)
                except Exception:
                    pass

            if isLaz:
                result["trimbleConnectLazResampleMethod"] = next(k for k, rb in self.tcLazMethodRadios.items() if rb.IsChecked)
                try:
                    result["trimbleConnectLazResampleSpacing"] = float(self.tcLazSpacingBox.Text)
                except Exception:
                    pass

            if (isOrtho or isLaz) and self.tcMaxSizeGbBox.Text.strip():
                try:
                    result["trimbleConnectMaxSizeGb"] = float(self.tcMaxSizeGbBox.Text)
                except Exception:
                    pass

        result["civilloEnabled"] = bool(self.civEnabledCheckbox.IsChecked)
        result["trimbleConnectEnabled"] = bool(self.tcEnabledCheckbox.IsChecked)
        result["label"] = propFileDisplayText + " -> " + " | ".join(targetLabels)
        self.result = result
        self.Close()

    def cancel_clicked(self, sender, e):
        self.result = None
        self.Close()

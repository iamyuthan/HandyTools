(function () {
  var w = window;
  var d = w["document"];

  // ---------- STATE ----------
  var rootDir = null;
  var sessionFolder = null;
  var pendingFolder = null;
  var sessionName = null;
  var sessionInitialized = false;
  var currentSessionKey = null;

  var cachedFirstUserMessage = null;

  var activeConversationFolder = null;
  var activeQuestionId = null;
  var conversationCounter = 0;

  var pendingFiles = [];

  var savedImages = new (w["Map"])();
  var savedHashes = new (w["Map"])();
  var savedFilesSet = new (w["Set"])();

  var DUPLICATE_WINDOW_MS = 60000;

  function log() {
    try {
      var args = w["Array"]["prototype"]["slice"]["call"](arguments);
      args["unshift"]("[TwinMind Saver v18]");
      w["console"]["log"]["apply"](w["console"], args);
    } catch (e) {}
  }

  // ---------- TEXT HELPERS ----------
  function sanitizeForFolder(text) {
    if (!text) return "untitled";
    return String(text)
      ["replace"](/[^a-zA-Z0-9\s_-]/g, "")
      ["replace"](/\s+/g, "_")
      ["replace"](/_+/g, "_")
      ["replace"](/^_|_$/g, "")
      ["slice"](0, 160);
  }

  function sanitizeFileName(name) {
    if (!name) return "file";
    return String(name)["replace"](/[<>:"\/\\|?*]/g, "_")["slice"](0, 100);
  }

  function first5WordsUnderscore(text) {
    if (!text) return "empty";
    var parts = String(text)["trim"]()["split"](/\s+/)["slice"](0, 5);
    var clean = [];
    for (var i = 0; i < parts["length"]; i++) {
      var w2 = parts[i]["replace"](/[^a-zA-Z0-9]/g, "");
      if (w2) clean["push"](w2);
    }
    return clean["length"] ? clean["join"]("_") : "message";
  }

  function first5WordsNoSpace(text) {
    if (!text) return "empty";
    var parts = String(text)["trim"]()["split"](/\s+/)["slice"](0, 5);
    var clean = [];
    for (var i = 0; i < parts["length"]; i++) {
      var w2 = parts[i]["replace"](/[^a-zA-Z0-9]/g, "");
      if (w2) clean["push"](w2);
    }
    return clean["length"] ? clean["join"]("") : "message";
  }

  function findExistingFirstUserMessage() {
    var sels = ['[data-role="user"]', ".user-message"];
    for (var i = 0; i < sels["length"]; i++) {
      try {
        var el = d["querySelector"](sels[i]);
        if (el && el["textContent"] && String(el["textContent"])["trim"]()) {
          return String(el["textContent"])["trim"]();
        }
      } catch (e) {}
    }
    return null;
  }

  function getStableFirstUserMessage() {
    if (cachedFirstUserMessage) return cachedFirstUserMessage;
    var m = findExistingFirstUserMessage();
    cachedFirstUserMessage = m || "session";
    return cachedFirstUserMessage;
  }

  function getExistingSessionIdForAttachments() {
    try {
      var ls = w["localStorage"];
      if (ls) {
        var cId = ls["getItem"]("currentConversationId");
        if (cId) {
          cId = String(cId)["replace"](/^["']|["']$/g, "");
          if (/^[a-f0-9-]{8,}/i["test"](cId)) return cId;
        }
      }
    } catch (e) {}
    try {
      var href = String(w["location"]["href"]);
      var m = href["match"](/[a-f0-9]{8}-[a-f0-9-]{4,}/i);
      if (m && m[0]) return m[0];
    } catch (e2) {}
    return "sess_" + w["Date"]["now"]()["toString"](36);
  }

  function getFallbackSessionId() {
    return "sess_" + w["Date"]["now"]()["toString"](36);
  }

  function getTimestamp() {
    return new (w["Date"])()["toISOString"]()["replace"](/[:.]/g, "-")["slice"](0, 19);
  }

  function getReadableTimestamp() {
    return new (w["Date"])()["toLocaleString"]();
  }

function getFolderTimeStamp() {
  var d = new (w["Date"])();
  var yyyy = d["getFullYear"]();
  var mm   = String(d["getMonth"]() + 1)["padStart"](2, "0");
  var dd   = String(d["getDate"]())["padStart"](2, "0");
  var hh   = String(d["getHours"]())["padStart"](2, "0");
  var mi   = String(d["getMinutes"]())["padStart"](2, "0");
  var ss   = String(d["getSeconds"]())["padStart"](2, "0");
  // 2026-02-20_01-01-17
  return yyyy + "-" + mm + "-" + dd + "_" + hh + "-" + mi + "-" + ss;
}

  // ---------- HASH / DUPLICATE ----------
  async function hashBlob(blob) {
    var buf = await blob["arrayBuffer"]();
    var hashBuf = await w["crypto"]["subtle"]["digest"]("SHA-256", buf);
    var arr = w["Array"]["from"](new (w["Uint8Array"])(hashBuf));
    return arr["map"](function (b) {
      return b["toString"](16)["padStart"](2, "0");
    })["join"]("")["slice"](0, 16);
  }

  async function isDuplicateImage(blob, width, height) {
    var now = w["Date"]["now"]();
    width = width || 0;
    height = height || 0;

    if (width && height) {
      var sizeBucket = w["Math"]["round"](blob["size"] / 10240) * 10;
      var fp = width + "x" + height + "_" + sizeBucket + "KB";
      var last = savedImages["get"](fp);
      if (last && now - last < DUPLICATE_WINDOW_MS) return true;
      savedImages["set"](fp, now);
    }

    var hash = await hashBlob(blob);
    var hLast = savedHashes["get"](hash);
    if (hLast && now - hLast < DUPLICATE_WINDOW_MS) return true;
    savedHashes["set"](hash, now);
    return false;
  }

  // ---------- ROOT + SESSION FOLDERS ----------
  async function ensureRootDir() {
    if (rootDir) return true;
    try {
      rootDir = await w["showDirectoryPicker"]({
        "mode": "readwrite",
        "startIn": "documents"
      });
      log("root folder chosen");
      return true;
    } catch (e) {
      log("root folder denied", e && e["message"]);
      return false;
    }
  }

  async function ensureSessionFolderEarly() {
    if (sessionFolder) return sessionFolder;
    var ok = await ensureRootDir();
    if (!ok) return null;

    var sessId = getExistingSessionIdForAttachments();
    var firstMsg = getStableFirstUserMessage();
    var first = first5WordsUnderscore(firstMsg || "session");
    var name = sanitizeForFolder(first + "_" + sanitizeForFolder(sessId));

    try {
      sessionFolder = await rootDir["getDirectoryHandle"](name, { "create": true });
      sessionName = name;
      sessionInitialized = true;
      currentSessionKey = sessId;
      log("session folder (early)", name);
      return sessionFolder;
    } catch (e) {
      log("session folder (early) error", e && e["message"]);
      return null;
    }
  }

  async function ensureSessionFolder(userMessage, sessionIdFromBody) {
    var ok = await ensureRootDir();
    if (!ok) return null;
    if (sessionFolder && sessionInitialized) return sessionFolder;

    var raw = sessionIdFromBody || getFallbackSessionId();
    var safeSessionId = sanitizeForFolder(raw);
    var firstMsg = getStableFirstUserMessage();
    var first = first5WordsUnderscore(firstMsg || "session");
    var name = sanitizeForFolder(first + "_" + safeSessionId);

    try {
      sessionFolder = await rootDir["getDirectoryHandle"](name, { "create": true });
      sessionName = name;
      sessionInitialized = true;
      currentSessionKey = raw;
      log("session folder (from body)", name);
      return sessionFolder;
    } catch (e) {
      log("session folder (body) error", e && e["message"]);
      return null;
    }
  }

  async function ensurePendingFolder() {
    var sess = await ensureSessionFolderEarly();
    if (!sess) return null;
    if (pendingFolder) return pendingFolder;
    try {
      pendingFolder = await sessionFolder["getDirectoryHandle"]("_pending", { "create": true });
      log("pending folder created");
      return pendingFolder;
    } catch (e) {
      log("pending folder error", e && e["message"]);
      return null;
    }
  }

  async function createConversationFolder(userMessage, questionIdRaw, sessionIdFromBody) {
    var parent = (await ensureSessionFolder(userMessage, sessionIdFromBody)) || sessionFolder;
    if (!parent) return null;

    var first5 = first5WordsNoSpace(userMessage);

    var qId = questionIdRaw && String(questionIdRaw)["trim"]();
    if (!qId) qId = "q_" + w["Date"]["now"]()["toString"](36);
    var safeQId = sanitizeForFolder(qId);

    var name = sanitizeForFolder(first5 + "_" + safeQId);

    try {
      var folder = await parent["getDirectoryHandle"](name, { "create": true });
      log("conversation folder", name);
      return {
        "folder": folder,
        "name": name,
        "questionId": safeQId,
        "rawQuestionId": qId,
        "meta": {}
      };
    } catch (e) {
      log("conversation folder error", e && e["message"]);
      return null;
    }
  }

  async function saveTextToFolder(folder, filename, content) {
    if (!folder) return false;
    try {
      var fh = await folder["getFileHandle"](filename, { "create": true });
      var wri = await fh["createWritable"]();
      await wri["write"](content);
      await wri["close"]();
      return true;
    } catch (e) {
      log("save text error", filename, e && e["message"]);
      return false;
    }
  }

  async function saveUserMessage(folder, userMessage, ids) {
    ids = ids || {};
    var lines = [];
    lines["push"]("USER");
    lines["push"]("Time: " + getReadableTimestamp());
    if (ids["questionId"]) lines["push"]("QuestionId: " + ids["questionId"]);
    if (ids["rawQuestionId"] && ids["rawQuestionId"] !== ids["questionId"]) {
      lines["push"]("RawQuestionId: " + ids["rawQuestionId"]);
    }
    if (ids["sessionId"]) lines["push"]("SessionId: " + ids["sessionId"]);
    if (ids["url"]) lines["push"]("URL: " + ids["url"]);
    if (ids["conversationId"]) lines["push"]("ConversationId: " + ids["conversationId"]);
    if (ids["model"]) lines["push"]("Model: " + ids["model"]);
    lines["push"]("");
    lines["push"](userMessage || "");
    return saveTextToFolder(folder, "01_user.txt", lines["join"]("\n"));
  }

  // ---------- FILE SAVE HELPERS ----------
  function getExtension(mime, name) {
    if (name) {
      var parts = String(name)["split"](".");
      if (parts["length"] > 1) return parts[parts["length"] - 1]["toLowerCase"]();
    }
    var map = {
      "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif", "image/webp": "webp",
      "application/pdf": "pdf", "text/plain": "txt", "text/html": "html",
      "text/css": "css", "text/javascript": "js", "text/markdown": "md",
      "text/csv": "csv", "application/json": "json",
      "application/msword": "doc",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
      "application/vnd.ms-excel": "xls",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
      "audio/mpeg": "mp3", "audio/wav": "wav", "video/mp4": "mp4", "video/webm": "webm"
    };
    return map[mime] || "bin";
  }

  async function saveFileToFolder(targetFolder, blob, originalName, prefix) {
    var ok = await ensureRootDir();
    if (!ok) return false;
    var parent = targetFolder || sessionFolder || rootDir;

    originalName = originalName || "";
    prefix = prefix || "File";

    var ext = getExtension(blob["type"], originalName);
    var base = originalName
      ? sanitizeFileName(originalName["replace"](/\.[^.]+$/, ""))
      : "file";

    var filename = prefix + "_" + getTimestamp() + "_" + base + "." + ext;

    try {
      var fh = await parent["getFileHandle"](filename, { "create": true });
      var wri = await fh["createWritable"]();
      await wri["write"](blob);
      await wri["close"]();
      log("saved file", filename);
      return filename;
    } catch (e) {
      log("file save error", e && e["message"]);
      return null;
    }
  }

  async function convertToHDPng(file) {
    return new (w["Promise"])(function (resolve) {
      var img = new (w["Image"])();
      img["onload"] = function () {
        var canvas = d["createElement"]("canvas");
        canvas["width"] = img["naturalWidth"];
        canvas["height"] = img["naturalHeight"];
        var ctx = canvas["getContext"]("2d");
        ctx["imageSmoothingEnabled"] = false;
        ctx["drawImage"](img, 0, 0);
        canvas["toBlob"](function (blob) {
          resolve({
            "blob": blob || file,
            "width": img["naturalWidth"],
            "height": img["naturalHeight"]
          });
        }, "image/png", 1.0);
      };
      img["onerror"] = function () {
        resolve({ "blob": file, "width": 0, "height": 0 });
      };
      img["src"] = w["URL"]["createObjectURL"](file);
    });
  }

  async function saveImageHDToFolder(targetFolder, blob, width, height, prefix) {
    prefix = prefix || "Image_HD";
    var ok = await ensureRootDir();
    if (!ok) return null;
    var parent = targetFolder || sessionFolder || rootDir;

    if (await isDuplicateImage(blob, width, height)) return null;

    var filename = prefix + "_" + getTimestamp() + ".png";
    try {
      var fh = await parent["getFileHandle"](filename, { "create": true });
      var wri = await fh["createWritable"]();
      await wri["write"](blob);
      await wri["close"]();
      log("saved image", filename);
      return filename;
    } catch (e) {
      log("image save error", e && e["message"]);
      return null;
    }
  }

  // ---------- PENDING → QUESTION ----------
  async function movePendingToQuestionFolder(questionFolder) {
    if (!pendingFolder || !questionFolder) return;
    try {
      for await (var entry of pendingFolder["values"]()) {
        if (entry["kind"] === "file") {
          var file = await entry["getFile"]();
          var fh = await questionFolder["getFileHandle"](entry["name"], { "create": true });
          var wri = await fh["createWritable"]();
          await wri["write"](file);
          await wri["close"]();
          log("moved from pending:", entry["name"]);
          await pendingFolder["removeEntry"](entry["name"]);
        }
      }
      log("pending folder cleared");
    } catch (e) {
      log("move pending error", e && e["message"]);
    }
  }

  // ---------- ATTACHMENT HANDLER ----------
  async function handleFileInput(files, source) {
    source = source || "Upload";
    for (var i = 0; i < files["length"]; i++) {
      var f = files[i];
      var isImage = f["type"] && f["type"]["indexOf"]("image/") === 0;

      pendingFiles["push"]({
        "blob": f,
        "name": f["name"],
        "source": source,
        "isImage": isImage
      });

      if (activeConversationFolder) {
        if (isImage) {
          var r = await convertToHDPng(f);
          await saveImageHDToFolder(activeConversationFolder, r["blob"], r["width"], r["height"], source + "_HD");
        } else {
          await saveFileToFolder(activeConversationFolder, f, f["name"], source);
        }
      } else {
        var pf = await ensurePendingFolder();
        if (!pf) continue;

        if (isImage) {
          var r2 = await convertToHDPng(f);
          await saveImageHDToFolder(pendingFolder, r2["blob"], r2["width"], r2["height"], source + "_HD");
        } else {
          await saveFileToFolder(pendingFolder, f, f["name"], source);
        }
      }
    }
  }

  // ---------- RENAME QUESTION FOLDER ONCE run_id KNOWN ----------
  async function renameConversationFolderIfNeeded(userMessage, convInfo, finalQId) {
    if (!sessionFolder || !convInfo || !convInfo["folder"]) return;
    if (!finalQId) return;

  var oldName = convInfo["name"] || "";

  // NEW: get stored time prefix, or fall back to "now"
  var meta = convInfo["meta"] || {};
  var timePrefix = meta["folderTime"] || getFolderTimeStamp();  // e.g. 20-02-2026_01-01-17

  var first5 = first5WordsNoSpace(userMessage);
  var safeNewQ = sanitizeForFolder(finalQId);

  // Build core name: Howmuchtimeisthe_a6476a00-...
  var coreName = first5 + "_" + safeNewQ;

  // Final: 20-02-2026_01-01-17_Howmuchtimeisthe_a6476a00-...
  var newName = sanitizeForFolder(timePrefix + "_" + coreName);

  if (!oldName || oldName === newName) return;

    try {
      var newFolder = await sessionFolder["getDirectoryHandle"](newName, { "create": true });

      for await (var entry of convInfo["folder"]["values"]()) {
        if (entry["kind"] === "file") {
          var file = await entry["getFile"]();
          var fh = await newFolder["getFileHandle"](entry["name"], { "create": true });
          var wri = await fh["createWritable"]();
          await wri["write"](file);
          await wri["close"]();
        }
      }

      try {
        await sessionFolder["removeEntry"](oldName, { "recursive": true });
      } catch (e2) {
        log("remove old folder failed", e2 && e2["message"]);
      }

      convInfo["folder"] = newFolder;
      convInfo["name"] = newName;
      log("renamed conversation folder", oldName, "->", newName);
    } catch (e) {
      log("rename folder error", e && e["message"]);
    }
  }

  // ---------- RESPONSE STREAM ----------
  async function processResponseStream(response, convInfo, userMessage) {
    if (!convInfo || !convInfo["folder"]) return;
    var folder = convInfo["folder"];

    var rawQId = convInfo["questionId"] || convInfo["rawQuestionId"] || "";
    var streamQId = "";

    var rawEvents = [];
    var thinking = "";
    var answer = "";

    try {
      var headers = response["headers"];
      var ct = "";
      if (headers && headers["get"]) ct = headers["get"]("content-type") || "";
      ct = String(ct)["toLowerCase"]();

      if (ct["indexOf"]("text/event-stream") !== -1 || ct["indexOf"]("stream") !== -1) {
        var reader = response["body"] && response["body"]["getReader"]
          ? response["body"]["getReader"]()
          : null;
        if (reader) {
          var decoder = new (w["TextDecoder"])();
          var buffer = "";
          while (true) {
            var res = await reader["read"]();
            if (res["done"]) break;
            buffer += decoder["decode"](res["value"], { "stream": true });
            var parts = buffer["split"]("\n");
            buffer = parts["pop"]() || "";
            for (var i = 0; i < parts["length"]; i++) {
              var line = parts[i];
              if (!String(line)["trim"]()) continue;
              rawEvents["push"](line);

              if (line["indexOf"]("data:") === 0) {
                var dataStr = String(line)["slice"](5)["trim"]();
                if (!dataStr || dataStr === "[DONE]") continue;

                if (!streamQId) {
                  var m = dataStr["match"](/["']run_id["']\s*:\s*["']([^"']+)["']/);
                  if (m && m[1]) {
                    streamQId = m[1];
                    log("run_id from stream", streamQId);
                  }
                }

                var obj = null;
                try { obj = w["JSON"]["parse"](dataStr); } catch (e) {
                  answer += dataStr;
                }

                if (obj) {
                  if (obj["thinking"]) thinking += obj["thinking"];
                  if (obj["content"]) answer += obj["content"];
                  if (obj["answer"]) answer += obj["answer"];
                  if (obj["text"]) answer += obj["text"];
                  if (obj["response"] && obj["response"]["content"]) answer += obj["response"]["content"];
                  if (obj["response"] && obj["response"]["thinking"]) thinking += obj["response"]["thinking"];
                  if (obj["choices"] && obj["choices"][0] && obj["choices"][0]["delta"]) {
                    var delta = obj["choices"][0]["delta"];
                    if (delta["content"]) answer += delta["content"];
                    if (delta["reasoning_content"]) thinking += delta["reasoning_content"];
                  }
                }
              }
            }
          }
          if (buffer && String(buffer)["trim"]()) rawEvents["push"](buffer);
        }
      } else {
        var text = await response["text"]();
        rawEvents["push"](text);

        if (!streamQId) {
          var m2 = text["match"](/["']run_id["']\s*:\s*["']([^"']+)["']/);
          if (m2 && m2[1]) {
            streamQId = m2[1];
            log("run_id from text", streamQId);
          }
        }

        var obj2 = null;
        try { obj2 = w["JSON"]["parse"](text); } catch (e) {}
        if (obj2) {
          if (obj2["thinking"]) thinking = obj2["thinking"];
          else if (obj2["reasoning"]) thinking = obj2["reasoning"];
          if (obj2["answer"]) answer = obj2["answer"];
          else if (obj2["content"]) answer = obj2["content"];
          else if (obj2["text"]) answer = obj2["text"];
          else if (obj2["response"] && obj2["response"]["content"]) answer = obj2["response"]["content"];
        } else {
          answer = text;
        }
      }
    } catch (e) {
      rawEvents["push"]("ERROR: " + (e && e["message"] ? e["message"] : ""));
    }

    if (rawEvents["length"]) {
      var content = [
        "EVENT STREAM",
        "Time: " + getReadableTimestamp(),
        "Count: " + rawEvents["length"],
        "",
        rawEvents["join"]("\n")
      ]["join"]("\n");
      await saveTextToFolder(folder, "02_stream.txt", content);
    }

    if (thinking) {
      var tContent = [
        "THINKING",
        "Time: " + getReadableTimestamp(),
        "",
        thinking
      ]["join"]("\n");
      await saveTextToFolder(folder, "03_thinking.txt", tContent);
    }

    if (answer) {
      var aContent = [
        "ANSWER",
        "Time: " + getReadableTimestamp(),
        "",
        answer
      ]["join"]("\n");
      await saveTextToFolder(folder, "04_answer.txt", aContent);
    }

    var finalQId = streamQId || rawQId;
    if (streamQId) {
      convInfo["questionId"] = streamQId;
    }

    // Rewrite [01_user.txt](https://01_user.txt) so QuestionId = finalQId
    try {
      var meta = convInfo["meta"] || {};
      await saveUserMessage(folder, userMessage, {
        "questionId": finalQId,
        "rawQuestionId": convInfo["rawQuestionId"] || rawQId,
        "sessionId": meta["sessionId"] || "",
        "url": meta["url"] || "",
        "conversationId": meta["conversationId"] || "",
        "model": meta["model"] || ""
      });
    } catch (e) {
      log("rewrite 01_user error", e && e["message"]);
    }

    var fullLines = [];
    fullLines["push"]("FULL CONVERSATION");
    fullLines["push"]("Time: " + getReadableTimestamp());
    fullLines["push"]("QuestionId: " + finalQId);
    fullLines["push"]("");
    fullLines["push"]("USER:");
    fullLines["push"](userMessage || "");
    fullLines["push"]("");
    if (thinking) {
      fullLines["push"]("THINKING:");
      fullLines["push"](thinking);
      fullLines["push"]("");
    }
    fullLines["push"]("ANSWER:");
    fullLines["push"](answer || "");
    await saveTextToFolder(folder, "05_full.txt", fullLines["join"]("\n"));

    await renameConversationFolderIfNeeded(userMessage, convInfo, finalQId);

    w["setTimeout"](function () {
      if (activeQuestionId === finalQId || activeQuestionId === rawQId) {
        activeConversationFolder = null;
        activeQuestionId = null;
      }
    }, 30000);
  }

  // ---------- FETCH INTERCEPTOR ----------
  var originalFetch = w["fetch"];

  w["fetch"] = async function (input, init) {
    init = init || {};
    var url = typeof input === "string" ? input : (input && input["url"]) || "";
    var method = (init["method"] || "GET")["toUpperCase"]();

    var looksLikeChat = false;
    if (method === "POST") {
      if (
        url["indexOf"]("twinmind") !== -1 ||
        url["indexOf"]("/chat") !== -1 ||
        url["indexOf"]("/conversation") !== -1 ||
        url["indexOf"]("/api/") !== -1 ||
        url["indexOf"]("/v1/") !== -1
      ) {
        looksLikeChat = true;
      }
    }

    var bodyStr = null;
    if (init["body"] && typeof init["body"] === "string") {
      bodyStr = init["body"];
    }

    var convInfo = null;
    var userMessageForThis = null;

    if (looksLikeChat && bodyStr) {
      var data = null;
      try { data = w["JSON"]["parse"](bodyStr); } catch (e) {}

      if (data) {
        var sessionId =
          data["session_id"] ||
          data["sessionId"] ||
          data["conversationSessionId"] ||
          data["conversationId"] ||
          data["conversation_id"] || "";

        var questionId =
          data["run_id"] || data["runId"] ||            // prefer run_id if in body
          data["questionId"] ||
          data["question_id"] ||
          data["id"] ||
          data["messageId"] || data["message_id"] ||
          data["requestId"] || data["request_id"] || "";

        var u =
          data["message"] ||
          data["query"] ||
          data["prompt"] ||
          data["content"] ||
          data["text"] ||
          "";

        if (!u && data["messages"] && data["messages"]["length"]) {
          var last = data["messages"][data["messages"]["length"] - 1];
          if (last && last["content"]) u = last["content"];
        }

        if (u && String(u)["length"] > 1) {
          convInfo = await createConversationFolder(u, questionId, sessionId);
          if (convInfo) {
            activeConversationFolder = convInfo["folder"];
            activeQuestionId = convInfo["questionId"];
            userMessageForThis = u;

            var meta = {
              "sessionId": sessionId,
              "url": url,
              "conversationId": data["conversationId"] || data["conversation_id"] || "",
              "model": data["model"] || "",
              "folderTime": getFolderTimeStamp()
            };
            convInfo["meta"] = meta;

            await saveUserMessage(convInfo["folder"], u, {
              "questionId": convInfo["questionId"],
              "rawQuestionId": convInfo["rawQuestionId"],
              "sessionId": meta["sessionId"],
              "url": meta["url"],
              "conversationId": meta["conversationId"],
              "model": meta["model"]
            });

            if (pendingFiles["length"]) {
              await movePendingToQuestionFolder(convInfo["folder"]);
              pendingFiles = [];
            }
          }
        }
      }
    }

    var response = await originalFetch["apply"](this, arguments);

    try {
      if (convInfo && userMessageForThis) {
        var clone = response["clone"]();
        processResponseStream(clone, convInfo, userMessageForThis);
      }
    } catch (e) {
      log("response clone error", e && e["message"]);
    }

    return response;
  };

  // ---------- FILE INPUT HOOKS ----------
  d["addEventListener"]("paste", function (e) {
    var cd = e["clipboardData"];
    if (!cd || !cd["items"]) return;
    var items = cd["items"];
    var files = [];
    for (var i = 0; i < items["length"]; i++) {
      var it = items[i];
      if (it && it["kind"] === "file") {
        var f = it["getAsFile"] && it["getAsFile"]();
        if (f) files["push"](f);
      }
    }
    if (files["length"]) handleFileInput(files, "Paste");
  }, true);

  d["addEventListener"]("change", function (e) {
    var t = e["target"];
    if (!t || t["type"] !== "file") return;
    var list = t["files"];
    if (!list || !list["length"]) return;
    var files = [];
    for (var i = 0; i < list["length"]; i++) files["push"](list[i]);
    handleFileInput(files, "Upload");
  }, true);

  d["addEventListener"]("drop", function (e) {
    var dt = e["dataTransfer"];
    if (!dt || !dt["files"] || !dt["files"]["length"]) return;
    var files = [];
    for (var i = 0; i < dt["files"]["length"]; i++) {
      files["push"](dt["files"][i]);
    }
    handleFileInput(files, "Drop");
  }, true);

  // ---------- INIT ----------
  (function init() {
    ensureRootDir()["then"](function (ok) {
      if (!ok) {
        log("no root folder chosen; saver idle");
      } else {
        log("ready; v18 (v11-style session + run_id rename + pending)");
      }
    });
  })();
})();

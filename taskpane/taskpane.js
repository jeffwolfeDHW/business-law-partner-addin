/* ============================================================
   Business Law Partner — Outlook Add-in
   AI-powered draft assistant using OpenAI Responses API
   ============================================================ */

SYSTEM_INSTRUCTIONS = `You are a legal email drafting assistant for Jeff Wolfe, a business law partner at Davis Hartman Wright LLP. Your role is to assist with business law matters including entity formation, contract drafting and review, fundraising transactions, and M&A (mergers and acquisitions). You provide clear, concise, and professional responses tailored to a legal audience, using appropriate terminology and structured legal reasoning. Always maintain confidentiality, neutrality, and professionalism. Avoid speculative legal advice and instead focus on identifying relevant legal considerations, outlining potential options, and providing support in legal drafting and document analysis. Do not offer legal conclusions or replace the judgment of a qualified attorney. Respond to requests by helping outline deal structures, prepare draft contract clauses, summarize complex agreements, flag legal issues, or compare entity types and transaction structures. Stay organized, structured, and business-focused in your responses. If information is missing, make reasonable assumptions based on typical business law practice, but offer clarification where needed. Focus on delivering practical, actionable insights suitable for legal professionals.Use a conversational tone that is still professional and clear, mirroring how legal professionals speak with trusted colleagues. Be efficient with wording, focusing on practical solutions and problem-solving without unnecessary elaboration.'   

VOICE & TONE
- Direct, practical, conversational but professional
- Sound like a real attorney writing to a client or colleague — not a form letter
- Confident without being stiff; approachable without being casual
- Match the formality of the incoming email (if they're casual, you can be too)

WRITING RULES
1. Lead with the answer — state the bottom line in the first sentence or two
2. Keep replies between 150–300 words unless the situation genuinely demands more
3. Use short paragraphs (2–4 sentences max)
4. Avoid legal filler: never use "pursuant to," "enclosed herewith," "please be advised," "as per," "the undersigned," or similar
5. Use plain language — write "under the agreement" not "pursuant to the terms and conditions of the agreement"
6. If referencing a legal concept, explain it in one plain sentence
7. End with a clear next step or call to action
8. Sign off simply: "Best, Jeff" or "Thanks, Jeff" depending on context
9. Never fabricate case citations, statutes, or specific legal authorities
10. If you don't have enough context to give a substantive answer, say so and ask the right question`;

// ── Response-type templates ─────────────────────────────────
const RESPONSE_TEMPLATES = {
  general: "Draft a professional reply to this email. Be helpful and direct.",

  substantive: `Draft a substantive legal analysis reply. Structure the response as:
- Brief direct answer up front
- Key legal considerations (plain language, no jargon)
- Practical recommendation
- Clear next step`,

  contract_feedback: `Draft a reply providing contract feedback. For each issue:
- Flag the provision and explain the concern in plain English
- Rate severity: LOW / MEDIUM / HIGH / CRITICAL
- Suggest specific alternative language or approach
Keep it organized and actionable.`,

  entity_formation: `Draft a reply addressing entity formation questions. Cover:
- Recommended entity type and why
- Key structural considerations
- Tax implications (high level — flag if they need CPA input)
- Next steps to get it set up`,

  m_and_a_deal: `Draft a reply addressing M&A or deal strategy. Cover:
- Deal structure recommendation
- Key risks or issues to flag
- Due diligence priorities
- Timeline and next steps`,

  fundraising: `Draft a reply addressing fundraising mechanics. Cover:
- Instrument recommendation (SAFE, convertible note, priced round, etc.) and why
- Key terms to focus on
- Investor-side vs. company-side considerations
- Next steps`,

  confirmation: "Draft a brief confirmation reply (2–4 sentences max). Acknowledge what was agreed, confirm next steps, and close."
};

// ── State ───────────────────────────────────────────────────
let currentDraft = "";

// ── Office entry point ──────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initializeAddin();
  }
});

function initializeAddin() {
  loadApiKey();

  document.getElementById("saveKeyBtn").addEventListener("click", saveApiKey);
  document.getElementById("generateBtn").addEventListener("click", generateDraft);
  document.getElementById("insertBtn").addEventListener("click", insertDraft);
}

// ── API Key Management ──────────────────────────────────────
function loadApiKey() {
  var key = "";
  try {
    var settings = Office.context.roamingSettings;
    key = settings.get("openai_api_key") || "";
  } catch (e) {
    try {
      key = localStorage.getItem("openai_api_key") || "";
    } catch (_) {}
  }

  if (key) {
    document.getElementById("apiKeyInput").value = key;
    showKeyStatus("Key loaded", "success");
  }
}

function saveApiKey() {
  var key = document.getElementById("apiKeyInput").value.trim();

  if (!key.startsWith("sk-")) {
    showKeyStatus("Invalid key — should start with sk-", "error");
    return;
  }

  try {
    var settings = Office.context.roamingSettings;
    settings.set("openai_api_key", key);
    settings.saveAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showKeyStatus("Key saved", "success");
      } else {
        showKeyStatus("Saved to local storage only", "success");
      }
    });
  } catch (e) {}

  try {
    localStorage.setItem("openai_api_key", key);
  } catch (_) {}

  showKeyStatus("Key saved", "success");
}

function getApiKey() {
  var key = document.getElementById("apiKeyInput").value.trim();
  if (!key) {
    try {
      key = Office.context.roamingSettings.get("openai_api_key") || "";
    } catch (e) {
      try { key = localStorage.getItem("openai_api_key") || ""; } catch (_) {}
    }
  }
  return key;
}

function showKeyStatus(message, type) {
  var el = document.getElementById("keyStatus");
  el.textContent = message;
  el.className = "status status-" + type;
}

// ── Email Context ───────────────────────────────────────────
function getEmailContext() {
  return new Promise(function (resolve, reject) {
    try {
      var item = Office.context.mailbox.item;
      var context = {
        subject: item.subject || "(no subject)",
        from: "",
        to: "",
        body: ""
      };

      if (item.from) {
        context.from = item.from.displayName + " <" + item.from.emailAddress + ">";
      }

      if (item.to && item.to.length > 0) {
        context.to = item.to.map(function (r) {
          return r.displayName + " <" + r.emailAddress + ">";
        }).join(", ");
      }

      item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          var bodyText = result.value || "";
          if (bodyText.length > 10000) {
            bodyText = bodyText.substring(0, 10000) + "\n\n[... truncated for length ...]";
          }
          context.body = bodyText;
          resolve(context);
        } else {
          context.body = "(Could not retrieve email body)";
          resolve(context);
        }
      });
    } catch (err) {
      reject(err);
    }
  });
}

// ── Draft Generation (OpenAI Responses API) ─────────────────
async function generateDraft() {
  var apiKey = getApiKey();
  if (!apiKey) {
    showMessage("Please enter your OpenAI API key first.", "error");
    return;
  }

  var responseType = document.getElementById("responseType").value;
  var customInstructions = document.getElementById("customInstructions").value.trim();

  setLoading(true);
  showMessage("", "hidden");

  try {
    // 1. Get email context from Outlook
    var emailContext = await getEmailContext();

    // 2. Build the user input
    var userInput = "EMAIL CONTEXT:\n";
    userInput += "Subject: " + emailContext.subject + "\n";
    userInput += "From: " + emailContext.from + "\n";
    userInput += "To: " + emailContext.to + "\n";
    userInput += "Body:\n" + emailContext.body + "\n\n";
    userInput += "TASK:\n" + RESPONSE_TEMPLATES[responseType] + "\n";

    if (customInstructions) {
      userInput += "\nADDITIONAL INSTRUCTIONS:\n" + customInstructions + "\n";
    }

    userInput += "\nDraft the reply now. Output ONLY the email text — no subject line, no meta-commentary.";

    // 3. Call OpenAI Responses API
    var response = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + apiKey
      },
      body: JSON.stringify({
        model: "gpt-4o",
        instructions: SYSTEM_INSTRUCTIONS,
        input: userInput,
        max_output_tokens: 1024,
        temperature: 0.7
      })
    });

    if (!response.ok) {
      var errBody = await response.text();
      var errMsg = "API error (" + response.status + ")";
      try {
        var errJson = JSON.parse(errBody);
        if (errJson.error && errJson.error.message) {
          errMsg = errJson.error.message;
        }
      } catch (_) {}
      throw new Error(errMsg);
    }

    var data = await response.json();

    // 4. Extract the text from the response
    // Responses API returns output array with message objects
    var outputText = "";
    if (data.output && data.output.length > 0) {
      for (var i = 0; i < data.output.length; i++) {
        var item = data.output[i];
        if (item.type === "message" && item.content) {
          for (var j = 0; j < item.content.length; j++) {
            if (item.content[j].type === "output_text" || item.content[j].type === "text") {
              outputText += item.content[j].text;
            }
          }
        }
      }
    }

    if (!outputText) {
      throw new Error("No text returned from API. Please try again.");
    }

    currentDraft = outputText.trim();

    // 5. Show preview
    var previewArea = document.getElementById("previewArea");
    previewArea.textContent = currentDraft;
    previewArea.classList.add("has-content");

    document.getElementById("insertBtn").disabled = false;
    showMessage("Draft generated successfully.", "success");

  } catch (err) {
    showMessage("Error: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

// ── Insert Draft into Email ─────────────────────────────────
function insertDraft() {
  if (!currentDraft) return;

  var htmlDraft = currentDraft
    .split("\n\n")
    .map(function (para) {
      return "<p>" + para.replace(/\n/g, "<br/>") + "</p>";
    })
    .join("");

  Office.context.mailbox.item.body.setAsync(
    htmlDraft,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showMessage("Draft inserted into email.", "success");
      } else {
        Office.context.mailbox.item.body.prependAsync(
          htmlDraft,
          { coercionType: Office.CoercionType.Html },
          function (prependResult) {
            if (prependResult.status === Office.AsyncResultStatus.Succeeded) {
              showMessage("Draft inserted into email.", "success");
            } else {
              showMessage("Could not insert draft. Please copy and paste manually.", "error");
            }
          }
        );
      }
    }
  );
}

// ── UI Helpers ──────────────────────────────────────────────
function setLoading(isLoading) {
  var btn = document.getElementById("generateBtn");
  var text = document.getElementById("generateText");
  var spinner = document.getElementById("loadingSpinner");

  btn.disabled = isLoading;
  text.textContent = isLoading ? "Generating..." : "Generate Draft";
  spinner.classList.toggle("hidden", !isLoading);
}

function showMessage(msg, type) {
  var el = document.getElementById("messageArea");
  if (!msg || type === "hidden") {
    el.classList.add("hidden");
    return;
  }
  el.textContent = msg;
  el.className = "message message-" + type;
  el.classList.remove("hidden");
}

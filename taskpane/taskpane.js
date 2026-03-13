/* ============================================================
   Business Law Partner — Outlook Add-in
   Core logic: Office.js integration + Claude API
   ============================================================ */

// ── System prompt (adapted from ChatGPT Business Law Partner GPT) ──

const SYSTEM_PROMPT = `You are drafting email replies on behalf of Jeff Wolfe, a business law partner at DHW Legal. Your role is to assist with business law matters including entity formation, contract drafting and review, fundraising transactions, and M&A (mergers and acquisitions).

Always maintain confidentiality, neutrality, and professionalism. Avoid speculative legal advice and instead focus on identifying relevant legal considerations, outlining potential options, and providing practical guidance. Do not offer legal conclusions or replace the judgment of a qualified attorney.

JEFF'S WRITING STYLE — match this closely:
- Greeting: Use the recipient's first name only, followed by a comma (e.g., "Miranda," or "John,"). Never use "Dear" or "Hi". If you cannot determine the first name, skip the greeting.
- Lead with the point: Get to the substance immediately. No preamble like "Thank you for reaching out" or "I hope this finds you well" unless the situation calls for warmth (e.g., acknowledging someone's patience).
- Be direct and honest: If something is outside scope or uncertain, say so plainly. Jeff writes things like "This isn't my area of expertise" rather than hedging.
- Brevity is paramount: Keep replies as short as possible while still being complete. Avoid filler, throat-clearing, and unnecessary background.
- Structure matches complexity: For simple replies, keep it to a few sentences. For complex analysis, use clear headers and organized sections. Let the content dictate the format.
- Practical and actionable: Always provide concrete next steps, options, or resources. End with an offer to help further when appropriate (e.g., "If helpful, I can draft three versions of this clause" or "Let me know if you'd like to discuss").
- Tone: Warm but efficient. Professional without being stiff. Like a trusted colleague at the next desk.
- Sign-off: Use "Jeff" for brief/casual replies, or "Regards,\\nJeff" for more formal ones. Never use "Best regards," "Sincerely," or "Warm regards."
- Keep it tight: Jeff does not over-explain. If two sentences will do, don't write four.

The draft should:
- Be ready to send with minimal editing
- Address all substantive points raised
- Include appropriate legal caveats where needed
- NOT include a subject line (this is a reply)
- NOT include an email signature block (Jeff's Outlook adds this automatically)`;

// ── Response type instructions ──

const RESPONSE_TYPES = {
  confirmation: "CONFIRMATION_PLACEHOLDER",
  substantive: "Read the email and identify every question and concern. Draft a brief, substantive reply that addresses each point with clear legal reasoning and actionable next steps. Keep it as short as possible — 2-5 sentences per point. Flag risks and recommend a course of action without over-explaining.",
  contract_feedback: "Review the contract terms referenced in this email. Flag only the most important concerns — provisions that are unusual, one-sided, or missing. Suggest specific fixes in 1-2 sentences each. Skip routine terms that are market-standard. Keep the overall reply concise.",
  entity_formation: "Give a brief entity formation recommendation. State the recommended entity type and the top 2-3 reasons why. If the client needs to provide more info before you can advise, say so in one sentence. Don't explain every entity type — just the recommended one.",
  general: "Draft a brief, professional reply addressing the key points in this email. Keep it to 2-5 sentences unless the complexity truly demands more. Get to the point immediately."
};

// ── Date helper for Confirmation replies ──

function getBusinessDaysLater(numDays) {
  var date = new Date();
  var added = 0;
  while (added < numDays) {
    date.setDate(date.getDate() + 1);
    var day = date.getDay();
    if (day !== 0 && day !== 6) {
      added++;
    }
  }
  var options = { weekday: "long", year: "numeric", month: "long", day: "numeric" };
  return date.toLocaleDateString("en-US", options);
}

// ── State ──

let generatedDraft = "";
let _inMemoryApiKey = "";  // Fallback when localStorage is blocked

// ── Initialize when Office is ready ──

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initializeUI();
  }
});

function initializeUI() {
  // Load saved API key status
  const savedKey = getApiKey();
  if (savedKey) {
    showKeyStatus("API key configured", "success");
    collapseApiKeySection();
  }

  // Event listeners
  document.getElementById("saveKeyBtn").addEventListener("click", saveApiKey);
  document.getElementById("generateBtn").addEventListener("click", handleGenerate);
  document.getElementById("insertBtn").addEventListener("click", handleInsert);
}

// ── API Key Management ──

function getApiKey() {
  // Try roamingSettings first (persists across Outlook restarts)
  try {
    var roaming = Office.context.roamingSettings;
    var saved = roaming.get("blp_claude_api_key");
    if (saved) return saved;
  } catch (e) {}
  // Fall back to localStorage
  try {
    var stored = localStorage.getItem("blp_claude_api_key");
    if (stored) return stored;
  } catch (e) {}
  return _inMemoryApiKey || "";
}
function saveApiKey() {
  var input = document.getElementById("apiKeyInput");
  var key = input.value.trim();

  if (!key) {
    showKeyStatus("Please enter an API key", "error");
    return;
  }

  if (!key.startsWith("sk-ant-")) {
    showKeyStatus("Key should start with sk-ant-...", "error");
    return;
  }

  // Always store in memory
  _inMemoryApiKey = key;

  // Save to roamingSettings (persists across Outlook restarts)
  try {
    var roaming = Office.context.roamingSettings;
    roaming.set("blp_claude_api_key", key);
    roaming.saveAsync(function(result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.log("roamingSettings save failed");
      }
    });
  } catch (e) {
    // Fall back to localStorage
    try { localStorage.setItem("blp_claude_api_key", key); } catch (e2) {}
  }

  input.value = "";
  showKeyStatus("API key saved", "success");
  collapseApiKeySection();
}
function showKeyStatus(message, type) {
  var el = document.getElementById("keyStatus");
  el.textContent = message;
  el.className = "status status-" + type;
}

function collapseApiKeySection() {
  var input = document.getElementById("apiKeyInput");
  input.placeholder = "Key saved — enter new key to replace";
}

// ── Email Context Retrieval ──

function getEmailContext() {
  return new Promise(function (resolve, reject) {
    try {
      var item = Office.context.mailbox.item;

      // Get subject
      var subject = item.subject || "(no subject)";

      // Get recipients
      var toRecipients = [];
      if (item.to && item.to.length) {
        for (var i = 0; i < item.to.length; i++) {
          toRecipients.push(item.to[i].emailAddress || item.to[i].displayName);
        }
      }

      // Check for attachments
      var attachmentNames = [];
      if (item.attachments && item.attachments.length) {
        for (var j = 0; j < item.attachments.length; j++) {
          attachmentNames.push(item.attachments[j].name);
        }
      }

      // Get body (async)
      item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          var body = result.value || "";

          // Truncate very long bodies to stay within token limits
          if (body.length > 10000) {
            body = body.substring(0, 10000) + "\n\n[... email truncated for length ...]";
          }

          resolve({
            subject: subject,
            to: toRecipients.join(", "),
            body: body,
            attachments: attachmentNames
          });
        } else {
          reject(new Error("Could not read email body: " + (result.error ? result.error.message : "Unknown error")));
        }
      });
    } catch (e) {
      reject(new Error("Could not access email: " + e.message));
    }
  });
}

// ── Claude API ──

function buildUserPrompt(emailContext, responseType, customInstructions) {
  var prompt = "I need to reply to the following email.\n\n";
  prompt += "Subject: " + emailContext.subject + "\n";
  prompt += "To: " + emailContext.to + "\n";
  if (emailContext.attachments && emailContext.attachments.length > 0) {
    prompt += "Attachments: " + emailContext.attachments.join(", ") + "\n";
    prompt += "(Note: I can see the attachment filenames but cannot read their contents directly. Work with whatever contract language or details are quoted or described in the email body.)\n";
  }
  prompt += "\nEmail content:\n---\n" + emailContext.body + "\n---\n\n";

  // Handle Confirmation type with calculated follow-up date
  if (responseType === "confirmation") {
    var followUpDate = getBusinessDaysLater(3);
    var typeInstruction = "Draft a brief, professional confirmation of receipt. Acknowledge the client's email and its general subject matter (without getting into substantive analysis). Let them know you are reviewing the matter and will provide a substantive response by " + followUpDate + ". Keep the tone warm but professional — this is a holding response, not a substantive one. Do not attempt to answer any legal questions or provide analysis. If the email mentions urgency or a deadline, acknowledge that as well.";
    prompt += "Task: " + typeInstruction + "\n";
  } else {
    var typeInstruction = RESPONSE_TYPES[responseType] || RESPONSE_TYPES.general;
    prompt += "Task: " + typeInstruction + "\n";
  }

  if (customInstructions && customInstructions.trim()) {
    prompt += "\nAdditional instructions: " + customInstructions.trim() + "\n";
  }

  prompt += "\nPlease draft the reply email body only (no subject line).";

  return prompt;
}

async function callClaudeAPI(userPrompt) {
  var apiKey = getApiKey();
  if (!apiKey) {
    throw new Error("No API key configured. Please enter your Claude API key above.");
  }

  // Debug: show key info so we can verify it's being retrieved
  var keyPreview = apiKey.substring(0, 10) + "..." + apiKey.substring(apiKey.length - 4);
  console.log("Using API key: " + keyPreview + " (length: " + apiKey.length + ")");

  var response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true"
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-5-20250929",
      max_tokens: 1024,
      system: SYSTEM_PROMPT,
      messages: [
        {
          role: "user",
          content: userPrompt
        }
      ]
    })
  });

  if (!response.ok) {
    var errorData;
    try {
      errorData = await response.json();
    } catch (e) {
      throw new Error("API request failed with status " + response.status);
    }
    var errorMsg = (errorData.error && errorData.error.message) ? errorData.error.message : "Unknown API error";
    // Include key preview and status in error for debugging
    throw new Error("Status " + response.status + ": " + errorMsg + " [key: " + keyPreview + ", len: " + apiKey.length + "]");
  }

  var data = await response.json();

  if (data.content && data.content.length > 0 && data.content[0].text) {
    return data.content[0].text;
  }

  throw new Error("Unexpected API response format");
}

// ── Generate Handler ──

async function handleGenerate() {
  var generateBtn = document.getElementById("generateBtn");
  var generateText = document.getElementById("generateText");
  var spinner = document.getElementById("loadingSpinner");
  var insertBtn = document.getElementById("insertBtn");

  // Disable button, show spinner
  generateBtn.disabled = true;
  generateText.textContent = "Generating...";
  spinner.classList.remove("hidden");
  hideMessage();

  try {
    // Get email context
    var emailContext = await getEmailContext();

    // Get selected options
    var responseType = document.getElementById("responseType").value;
    var customInstructions = document.getElementById("customInstructions").value;

    // Build prompt and call API
    var userPrompt = buildUserPrompt(emailContext, responseType, customInstructions);
    var draft = await callClaudeAPI(userPrompt);

    // Display preview
    generatedDraft = draft;
    displayPreview(draft);

    // Enable insert button
    insertBtn.disabled = false;

    showMessage("Draft generated successfully", "success");
  } catch (error) {
    showMessage(error.message, "error");
    generatedDraft = "";
    insertBtn.disabled = true;
  } finally {
    generateBtn.disabled = false;
    generateText.textContent = "Generate Draft";
    spinner.classList.add("hidden");
  }
}

// ── Insert into Email ──

function handleInsert() {
  if (!generatedDraft) {
    showMessage("No draft to insert", "error");
    return;
  }

  try {
    var item = Office.context.mailbox.item;

    // Convert plain text to HTML that matches Outlook's default compose style
    // Use <br> for line breaks and <br><br> for paragraph breaks (single-spaced)
    var htmlDraft = generatedDraft
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\n\n/g, "<br><br>")
      .replace(/\n/g, "<br>");

    // Wrap in a div with Outlook-matching font style (inherits from compose window)
    htmlDraft = '<div style="font-family: Calibri, sans-serif; font-size: 11pt; margin: 0; padding: 0;">' + htmlDraft + '</div><br>';

    // Prepend draft to existing body (keeps the quoted reply thread)
    item.body.prependAsync(
      htmlDraft,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showMessage("Draft inserted into email", "success");
        } else {
          showMessage("Failed to insert: " + (result.error ? result.error.message : "Unknown error"), "error");
        }
      }
    );
  } catch (e) {
    showMessage("Insert failed: " + e.message, "error");
  }
}
function displayPreview(text) {
  var previewArea = document.getElementById("previewArea");
  previewArea.textContent = text;
  previewArea.classList.add("has-content");
}

function showMessage(text, type) {
  var el = document.getElementById("messageArea");
  el.textContent = text;
  el.className = "message message-" + type;
  el.classList.remove("hidden");

  if (type === "success") {
    setTimeout(function () {
      el.classList.add("hidden");
    }, 4000);
  }
}

function hideMessage() {
  var el = document.getElementById("messageArea");
  el.classList.add("hidden");
}

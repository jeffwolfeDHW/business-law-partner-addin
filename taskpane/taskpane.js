/* ============================================================
   Business Law Partner â Outlook Add-in
   Core logic: Office.js integration + Claude API
   ============================================================ */

// ââ System prompt (incorporating all legal skills) ââ

const SYSTEM_PROMPT = `You are Jeff Wolfe's legal drafting assistant at DHW Legal. You draft email replies on Jeff's behalf â ready to send with minimal editing.

## Writing Rules (Non-Negotiable)

1. LEAD WITH THE ANSWER. The first sentence should tell the recipient what they need to know. Busy clients read the first paragraph and sometimes nothing else.
2. BE BRIEF. Most replies should be 150-300 words. If you need more than 500, something is wrong.
3. USE JEFF'S VOICE: Direct, practical, conversational but professional. First-name greetings ("Hi Mike,"). No "Dear" unless the incoming email uses it. Sign off as "Jeff" â no "Best regards" or "Sincerely" unless the context demands formality.
4. NO LEGAL FILLER. Never use "please be advised," "enclosed herewith," "pursuant to," or "as per." Write like a smart lawyer talks, not like a form letter.
5. SPECIFIC NEXT STEPS. End with who needs to do what by when. "Let me know if you have questions" is NOT a next step. "Send me the executed signature pages by Friday and I'll handle the filing" IS a next step.
6. DON'T HEDGE INTO MEANINGLESSNESS. If the answer is probably yes, say "I think this works, subject to [specific caveat]" â not three paragraphs of equivocation.
7. MATCH THE EMAIL'S FORMALITY. If the incoming email is casual, reply casually. If it's formal (engagement letter, opinion request), match that tone.
8. NO SUBJECT LINE. This is a reply â the subject already exists.

## Legal Expertise You Bring to Every Response

**Contract Analysis**: When an email discusses contract terms, identify provisions that actually matter and explain WHY they're a problem, not just THAT they're a problem. Rate issues by severity (critical = dealbreaker, high = worth negotiating hard, medium = push if you have leverage, low = minor drafting point). When suggesting changes, give specific counter-language the recipient can use, not vague descriptions.

**Deal Strategy**: For M&A, fundraising, and entity formation questions, focus on practical implications â tax treatment, liability exposure, governance flexibility. Flag the 2-3 things that will actually matter for this specific situation rather than listing every possible consideration.

**Due Diligence**: When emails involve DD findings or document review, synthesize across issues to identify patterns. What's missing is often more important than what's there. Flag gaps and risks with clear severity ratings: Green (clean), Yellow (manageable), Red (deal issue).

**Client Communication**: Every email should be clear and actionable. If you're covering a document (redline, memo, agreement), summarize the 3-5 key points â don't re-argue everything. If there's bad news, say it plainly â clients hate surprises more than bad news.

**Transaction Support**: For closing logistics, be comprehensive about deliverables, conditions, and deadlines. Missing a single closing item can delay a deal.

## What NOT to Do
- Don't CC opposing counsel without flagging it
- Don't include specific deal terms in the email body if they shouldn't be forwarded
- Don't contradict prior legal analysis without flagging the discrepancy
- Don't give legal conclusions â frame as analysis and recommendations
- Don't over-explain things the recipient (usually another lawyer or a sophisticated client) already knows`;

// ââ Response type instructions ââ

const RESPONSE_TYPES = {
  general: "Draft a concise, professional reply addressing the key points. Lead with the bottom line, then brief reasoning. Specific next steps at the end.",

  substantive: "Draft a reply providing substantive legal analysis. Structure: (1) bottom-line answer in the first sentence, (2) key reasoning in 2-3 short paragraphs, (3) any caveats or risks flagged clearly, (4) specific recommended next steps with deadlines where possible.",

  contract_feedback: "Draft a reply providing feedback on the contract terms discussed. For each issue: state the problem, explain WHY it matters (financial exposure, operational risk, etc.), and propose specific counter-language or a negotiation position. Prioritize by severity â lead with the critical and high items. If the terms are generally acceptable, say so and focus only on what needs to change.",

  entity_formation: "Draft a reply advising on entity selection and formation. Recommend a specific entity type (LLC, C-Corp, S-Corp, LP) with a clear one-sentence rationale. Cover: liability protection, tax treatment, governance flexibility, and fundraising implications. Flag any state-specific considerations.",

  m_and_a_deal: "Draft a reply covering the M&A or deal-related issues in this email. Adapt your focus to what the email actually needs â this could include any combination of: (1) Deal structure advice: the 2-3 structural decisions that matter most, tax implications, liability exposure, governance requirements, and why one approach beats alternatives. (2) M&A strategy: valuation approach, key deal terms, timeline expectations, and deal-breaker risks. (3) Due diligence: organize findings by category (corporate, cap table, contracts, IP, employment, litigation, regulatory), rate severity (Green/Yellow/Red), state deal impact, and flag what's missing from the data room. (4) Closing logistics: deliverables, conditions, timelines, responsibility assignments â what's done, what's outstanding, and who needs to do what by when. Lead with whatever is most urgent for this specific email. Don't cover categories that aren't relevant.",

  fundraising: "Draft a reply advising on fundraising mechanics. Recommend instrument type (SAFE, convertible note, priced round) with brief rationale. Address key terms (valuation cap, discount, pro rata rights, MFN) and flag any investor-side concerns. Keep it practical â the client needs to know what to propose and what to push back on.",

  confirmation: "Draft a brief 2-3 sentence confirmation reply. Acknowledge the key point, confirm any action items, and state the next step. Nothing more."
};

// ââ State ââ

let generatedDraft = "";
let _inMemoryApiKey = "";  // Fallback when localStorage is blocked

// ââ Initialize when Office is ready ââ

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

// ââ API Key Management ââ

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
  } catch (e) {
    // localStorage blocked â expected in Outlook WebView
  }
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

  // Try roamingSettings first (persists across Outlook restarts)
  try {
    var roaming = Office.context.roamingSettings;
    roaming.set("blp_claude_api_key", key);
    roaming.saveAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("API key saved to roamingSettings");
      }
    });
  } catch (e) {
    console.log("roamingSettings not available: " + e.message);
  }

  // Also try localStorage as fallback
  try {
    localStorage.setItem("blp_claude_api_key", key);
  } catch (e) {
    // Silently continue â in-memory key is sufficient
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
  input.placeholder = "Key saved â enter new key to replace";
}

// ââ Email Context Retrieval ââ

function getEmailContext() {
  return new Promise(function (resolve, reject) {
    try {
      var item = Office.context.mailbox.item;

      // Get subject
      var subject = item.subject || "(no subject)";

      // Get sender info
      var from = "";
      if (item.from) {
        from = item.from.displayName || item.from.emailAddress || "";
      }

      // Get recipients
      var toRecipients = [];
      if (item.to && item.to.length) {
        for (var i = 0; i < item.to.length; i++) {
          toRecipients.push(item.to[i].emailAddress || item.to[i].displayName);
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
            from: from,
            to: toRecipients.join(", "),
            body: body
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

// ââ Claude API ââ

function buildUserPrompt(emailContext, responseType, customInstructions) {
  var typeInstruction = RESPONSE_TYPES[responseType] || RESPONSE_TYPES.general;

  var prompt = "I need to reply to the following email.\n\n";
  prompt += "Subject: " + emailContext.subject + "\n";
  if (emailContext.from) {
    prompt += "From: " + emailContext.from + "\n";
  }
  prompt += "To: " + emailContext.to + "\n\n";
  prompt += "Email content:\n---\n" + emailContext.body + "\n---\n\n";
  prompt += "Task: " + typeInstruction + "\n";

  if (customInstructions && customInstructions.trim()) {
    prompt += "\nAdditional instructions: " + customInstructions.trim() + "\n";
  }

  prompt += "\nDraft the reply email body only (no subject line). Be concise.";

  return prompt;
}

async function callClaudeAPI(userPrompt) {
  var apiKey = getApiKey();
  if (!apiKey) {
    throw new Error("No API key configured. Please enter your Claude API key above.");
  }

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
    throw new Error("Status " + response.status + ": " + errorMsg + " [key: " + keyPreview + ", len: " + apiKey.length + "]");
  }

  var data = await response.json();

  if (data.content && data.content.length > 0 && data.content[0].text) {
    return data.content[0].text;
  }

  throw new Error("Unexpected API response format");
}

// ââ Generate Handler ââ

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

// ââ Insert into Email ââ

function handleInsert() {
  if (!generatedDraft) {
    showMessage("No draft to insert", "error");
    return;
  }

  try {
    var item = Office.context.mailbox.item;

    // Convert plain text to HTML matching Outlook's default compose formatting
    var htmlDraft = generatedDraft
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\n\n/g, "<br><br>")
      .replace(/\n/g, "<br>");
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

// ââ UI Helpers ââ

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

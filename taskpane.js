(() => {
  const $ = (id) => document.getElementById(id);
  let currentEmail = "";

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (ch) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    })[ch]);
  }



function buildSignature(name) {
  const cfg = window.BRIDMET_SIGNATURE_CONFIG;
  const safeName = escapeHtml(name.trim());
  const imageUrl = "https://reecehowebm.github.io/outlook-signatures/bsi-certification.png";

  const certLines = cfg.certificationLines
    .map(line => escapeHtml(line))
    .join("<br>");

  return `
<table role="presentation" cellpadding="0" cellspacing="0" border="0"
  style="border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:inherit;">
  <tr>
    <td colspan="2" style="padding:0 0 6px 0;font-size:14px;">
      Kind Regards
    </td>
  </tr>
  <tr>
    <td valign="top" style="padding:0 14px 0 0;vertical-align:top;">
      <div style="margin:0 0 3px 0;font-size:14px;font-weight:bold;">
        ${safeName}
      </div>
      <div style="margin:0 0 5px 0;font-size:14px;font-weight:bold;">
        BRIDMET PURCHASING TEAM
      </div>
      <div style="margin:0;font-size:14px;font-weight:bold;">
        TEL: ${escapeHtml(cfg.telephone)}
      </div>
      <div style="margin:0;font-size:14px;">
        <a href="${escapeHtml(cfg.website)}" style="color:inherit;text-decoration:underline;">${escapeHtml(cfg.websiteLabel)}</a> | GORE
      </div>
      <div style="margin:0;font-size:14px;font-weight:bold;">
        ${escapeHtml(cfg.locationLine1)}
      </div>
      <div style="margin:0 0 6px 0;font-size:14px;font-weight:bold;">
        ${escapeHtml(cfg.locationLine2)}
      </div>
     <div style="margin:0;font-size:12px;line-height:16px;">
  ${certLines}
</div>
    </td>
    <td valign="top" style="padding:0;vertical-align:top;">
      <img src="${escapeHtml(imageUrl)}"
        alt="BSI ISO certification"
        width="220"
        style="display:block;width:220px;height:auto;border:0;">
    </td>
  </tr>
</table>
<br>`;
}



  function setStatus(message, isError = false) {
    $("status").textContent = message;
    $("status").style.color = isError ? "#b91c1c" : "#14532d";
  }

  function initialise() {
    const office = window.Office;
    if (!office || !office.context || !office.context.mailbox) {
      $("matchNote").textContent = "Open this pane from Outlook while composing an email.";
      return;
    }

    currentEmail = (office.context.mailbox.userProfile.emailAddress || "").trim().toLowerCase();
    const matched = (window.BRIDMET_SIGNATURE_CONFIG.users || []).find(
      u => String(u.email).trim().toLowerCase() === currentEmail
    );

    $("employeeName").disabled = false;
    $("insertButton").disabled = false;

    if (matched) {
      $("employeeName").value = matched.name;
      $("matchNote").textContent = `Matched your Outlook account: ${currentEmail}`;
    } else {
      const displayName = office.context.mailbox.userProfile.displayName || "";
      $("employeeName").value = displayName;
      $("matchNote").textContent = currentEmail
        ? `Your Outlook account is ${currentEmail}. Check or enter your name below.`
        : "Your Outlook name wasn't available. Enter your name below.";
    }

    $("insertButton").addEventListener("click", insertSignature);
  }

  function insertSignature() {
    const name = $("employeeName").value.trim();
    if (!name) {
      setStatus("Enter your name before inserting the signature.", true);
      $("employeeName").focus();
      return;
    }

    const item = Office.context.mailbox.item;
    if (!item || !item.body || typeof item.body.setSelectedDataAsync !== "function") {
      setStatus("This Outlook compose context doesn't support inserting HTML this way. Check that you're composing a message in a supported Outlook client.", true);
      return;
    }

    $("insertButton").disabled = true;
    setStatus("Inserting signature…");

    item.body.setSelectedDataAsync(
      buildSignature(name),
      { coercionType: Office.CoercionType.Html },
      (result) => {
        $("insertButton").disabled = false;
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus("Signature inserted. Check the layout and recipient view before sending.");
        } else {
          setStatus(`Couldn't insert the signature: ${result.error ? result.error.message : "Unknown error"}`, true);
        }
      }
    );
  }

  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) initialise();
    else $("matchNote").textContent = "This add-in must be opened in Outlook.";
  });
})();

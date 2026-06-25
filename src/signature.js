// CDN URLs for images - UPDATE THESE with your actual CDN
const CDN_BASE = "https://cdn.isg-one.com/Email%20Images";
const CC_BADGE_URL = `${CDN_BASE}/ISG_image017.png`;
const ISG_LOGO_URL = `${CDN_BASE}/ISG_image005_Trans.png`;

// Entry point - fires when user creates new email
function onNewMessageCompose(event) {
    console.log("New compose event triggered");

    Office.onReady(() => {
        console.log("Office ready");

        getUserProfile()
            .then(user => {
                console.log("Building signature for:", user.displayName);

                const signatureHtml = buildSignatureHtml(user);

                // Set the signature in the compose body
                Office.context.mailbox.item.body.setSignatureAsync(
                    signatureHtml,
                    { coercionType: Office.CoercionType.Html },
                    (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("✓ Signature set successfully");
                        } else {
                            console.error("✗ Signature failed:", result.error.message);
                        }
                        event.completed();
                    }
                );
            })
            .catch(error => {
                console.error("Error in onNewMessageCompose:", error);
                event.completed();
            });
    });
}

// Builds the signature HTML
function buildSignatureHtml(user) {
    console.log("Comments field:", user.comments);

    const designation = getDesignation(user.comments);
    const badgeUrl = getDesignationBadgeUrl(designation);
    const isCc = isChairmansClub(user.comments);
    const country = getCountry(user.comments);
    const busUnit = getBusUnit(user.comments);

    console.log("Designation:", designation);
    console.log("Badge URL:", badgeUrl);
    console.log("Chairman's Club:", isCc);
    console.log("Country:", country);
    console.log("Bus Unit:", busUnit);


    // Build badge cell
    let badgeCell = "";
    if (badgeUrl) {
        // 2026 badge - use image from CDN
        badgeCell = `<td style="padding-left:4px;">
            <img src="${badgeUrl}" 
                 alt="${designation}" 
                 title="${designation}" 
                 width="115" height="17" 
                 style="vertical-align:middle;display:block;"/>
        </td>`;
    } else if (designation) {
        // Non-2026 designation - show as text
        badgeCell = `<td style="padding-left:4px;">
            <span style="font-size:10pt;color:#767171;">
                <b>${designation}</b>
            </span>
        </td>`;
    }

    // Build Chairman's Club badge cell
    let ccBadgeCell = "";
    if (isCc) {
        ccBadgeCell = `<td style="padding-left:16px;padding-right:2px;">
            <img src="${CC_BADGE_URL}" 
                 alt="Chairman's Club" 
                 title="Chairman's Club" 
                 width="29" height="21" 
                 style="vertical-align:middle;display:block;"/>
        </td>`;
    }

    // Build optional rows
    let titleRow = "";
    if (user.jobTitle) {
        titleRow = `<tr><td colspan="3" style="font-size:10pt;color:#767171;padding-top:2px;">
            ${user.jobTitle}
        </td></tr>`;
    }

    let phoneRow = "";
    if (user.businessPhone) {
        phoneRow = `<tr><td colspan="3" style="font-size:10pt;color:#767171;padding-top:2px;">
            ${user.businessPhone} office
        </td></tr>`;
    }

    let mobileRow = "";
    if (user.mobilePhone) {
        mobileRow = `<tr><td colspan="3" style="font-size:10pt;color:#767171;padding-top:2px;">
            ${user.mobilePhone} mobile
        </td></tr>`;
    }

    // Country-specific legal footer (Germany / Switzerland)
    const countryFooter = getCountryFooterHtml(country);

    // General ISG description shown to all users
    const isgDescription = `
        <p style="margin:4px 0;">
            <span style="font-size:10pt;color:#767171;">
                <b>Information Services Group</b>
            </span>
            <span style="font-size:10pt;color:#767171;">
                &nbsp;is a global AI-centered technology research and advisory firm.
            </span>
        </p>`;

    // General confidentiality note shown to all users
    const confidentialityNote = `
        <p style="margin:4px 0;">
            <span style="font-size:8pt;color:#767171;">
                Confidentiality note: The above email contains information that is 
                confidential and/or privileged. The information is for the use of the 
                individual or entity originally intended. If you are not the intended 
                recipient, please contact the sender and delete the material from any 
                computer. Please be aware that any disclosure, copying, distribution 
                or use of this information is prohibited.
            </span>
        </p>`;

    // Build complete signature HTML
    const html = `
    <table cellpadding="0" cellspacing="0" border="0" style="font-family:Calibri,Arial,sans-serif;">
        <tr>
            <td style="padding-bottom:8px;">
                <span style="font-size:12pt;font-weight:bold;color:#1F3864;">
                    ${user.displayName}
                </span>
            </td>
            ${ccBadgeCell}
            ${badgeCell}
        </tr>
        ${titleRow}
        ${phoneRow}
        ${mobileRow}
        <tr>
            <td colspan="3" style="padding-top:8px;">
                <a href="https://www.isg-one.com" target="_blank" rel="noopener">
                    <img src="${ISG_LOGO_URL}" alt="ISG" width="150" height="50"/>
                </a>
            </td>
        </tr>
        <!-- Country-specific legal footer (Germany / Switzerland only) -->
        ${countryFooter ? `
        <tr>
            <td colspan="3">
                ${countryFooter}
            </td>
        </tr>` : ""}

        <!-- ISG description + confidentiality note (all users) -->
        <tr>
            <td colspan="3">
                ${isgDescription}
                ${confidentialityNote}
            </td>
        </tr>
    </table>`;

    return html;
}
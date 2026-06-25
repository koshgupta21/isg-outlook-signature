// ─────────────────────────────────────────────
// CDN URLs — update with your actual CDN
// ─────────────────────────────────────────────

const CDN_BASE = "https://your-cdn.isg-one.com/signatures";
const CC_BADGE_URL = `${CDN_BASE}/ISG_image017.png`;
const ISG_LOGO_URL = `${CDN_BASE}/ISG_logo.png`;

// ─────────────────────────────────────────────
// ENTRY POINT
// Fired by OnNewMessageCompose event in manifest
// ─────────────────────────────────────────────

function onNewMessageCompose(event) {
    console.log("onNewMessageCompose fired");

    Office.onReady(() => {
        console.log("Office ready");

        getUserProfile()
            .then(user => {
                const html = buildSignatureHtml(user);
                Office.context.mailbox.item.body.setSignatureAsync(
                    html,
                    { coercionType: Office.CoercionType.Html },
                    result => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("✓ Signature set successfully");
                        } else {
                            console.error("✗ setSignatureAsync failed:", result.error.message);
                        }
                        event.completed();
                    }
                );
            })
            .catch(error => {
                console.error("Error:", error);
                event.completed();
            });
    });
}

// ─────────────────────────────────────────────
// BUILD SIGNATURE HTML
// ─────────────────────────────────────────────

function buildSignatureHtml(user) {

    // ── Get all values ──────────────────────
    const designation = getDesignation(user.comments);
    const badgeUrl = getDesignationBadgeUrl(designation);
    const isCc = isChairmansClub(user.comments);
    const country = getCountry(user.comments);
    const busUnit = getBusUnit(user.comments);

    console.log("displayName :", user.displayName);
    console.log("designation :", designation);
    console.log("badgeUrl    :", badgeUrl);
    console.log("isCc        :", isCc);
    console.log("country     :", country);
    console.log("busUnit     :", busUnit);

    // ── Badge cell ──────────────────────────
    let badgeCell = "";
    if (badgeUrl) {
        badgeCell = `
            <td style="padding-left:8px;vertical-align:middle;">
                <img src="${badgeUrl}"
                     alt="${designation}"
                     title="${designation}"
                     width="115" height="17"/>
            </td>`;
    } else if (designation) {
        badgeCell = `
            <td style="padding-left:8px;vertical-align:middle;">
                <span style="font-size:10pt;color:#767171;
                             font-family:Calibri,Arial,sans-serif;">
                    <b>${designation}</b>
                </span>
            </td>`;
    }

    // ── Chairman's Club cell ─────────────────
    const ccCell = isCc ? `
        <td style="padding-left:8px;vertical-align:middle;">
            <img src="${CC_BADGE_URL}"
                 alt="Chairman's Club"
                 title="Chairman's Club"
                 width="29" height="21"/>
        </td>` : "";

    // ── Optional rows ────────────────────────
    const titleRow = user.jobTitle ? `
        <tr>
            <td colspan="3"
                style="font-size:10pt;color:#767171;padding-top:2px;
                       font-family:Calibri,Arial,sans-serif;">
                ${user.jobTitle}
            </td>
        </tr>` : "";

    const officePhoneRow = user.businessPhone ? `
        <tr>
            <td colspan="3"
                style="font-size:10pt;color:#767171;padding-top:2px;
                       font-family:Calibri,Arial,sans-serif;">
                ${user.businessPhone} office
            </td>
        </tr>` : "";

    const mobileRow = user.mobilePhone ? `
        <tr>
            <td colspan="3"
                style="font-size:10pt;color:#767171;padding-top:2px;
                       font-family:Calibri,Arial,sans-serif;">
                ${user.mobilePhone} mobile
            </td>
        </tr>` : "";

    // ── Country legal footer ─────────────────
    // Only appears for Germany and Switzerland
    const countryFooter = getCountryFooterHtml(country);

    const countryFooterRow = countryFooter ? `
        <tr>
            <td colspan="3" style="padding-top:8px;">
                ${countryFooter}
            </td>
        </tr>` : "";

    // ── ISG description (all users) ──────────
    const isgDescriptionRow = `
        <tr>
            <td colspan="3" style="padding-top:8px;">
                <p style="margin:0;">
                    <span style="font-size:10pt;color:#767171;
                                 font-family:Calibri,Arial,sans-serif;">
                        <b>Information Services Group</b>
                    </span>
                    <span style="font-size:10pt;color:#767171;
                                 font-family:Calibri,Arial,sans-serif;">
                        &nbsp;is a global AI-centered technology research and advisory firm.
                    </span>
                </p>
            </td>
        </tr>`;

    // ── Confidentiality note (all users) ─────
    const confidentialityRow = `
        <tr>
            <td colspan="3" style="padding-top:4px;">
                <p style="margin:0;">
                    <span style="font-size:8pt;color:#767171;
                                 font-family:Calibri,Arial,sans-serif;">
                        Confidentiality note: The above email contains information that is
                        confidential and/or privileged. The information is for the use of
                        the individual or entity originally intended. If you are not the
                        intended recipient, please contact the sender and delete the
                        material from any computer. Please be aware that any disclosure,
                        copying, distribution or use of this information is prohibited.
                    </span>
                </p>
            </td>
        </tr>`;

    // ── Assemble complete signature ──────────
    return `
    <table cellpadding="0" cellspacing="0" border="0"
           style="font-family:Calibri,Arial,sans-serif;">

        <!-- Row 1: Name + Chairman's Club + Badge -->
        <tr>
            <td style="vertical-align:middle;padding-bottom:4px;">
                <span style="font-size:12pt;font-weight:bold;color:#1F3864;
                             font-family:Calibri,Arial,sans-serif;">
                    ${user.displayName}
                </span>
            </td>
            ${ccCell}
            ${badgeCell}
        </tr>

        <!-- Row 2: Job title -->
        ${titleRow}

        <!-- Row 3: Office phone -->
        ${officePhoneRow}

        <!-- Row 4: Mobile -->
        ${mobileRow}

        <!-- Row 5: ISG Logo -->
        <tr>
            <td colspan="3" style="padding-top:8px;padding-bottom:8px;">
                <img src="${ISG_LOGO_URL}" alt="ISG" width="120" height="40"/>
            </td>
        </tr>

        <!-- Row 6: Country footer (Germany / Switzerland only) -->
        ${countryFooterRow}

        <!-- Row 7: ISG description (all users) -->
        ${isgDescriptionRow}

        <!-- Row 8: Confidentiality note (all users) -->
        ${confidentialityRow}

    </table>`;
}
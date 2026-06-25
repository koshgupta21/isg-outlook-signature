// Port of GetDesignation() from C# to JavaScript
function getDesignation(comments) {
    if (!comments) return "";

    const c = comments.toLowerCase();

    // Check for codes in order
    if (c.includes("daiadv"))   return "AI Certified | Advanced";
    if (c.includes("exa") ||
        c.includes("exadap") ||
        c.includes("exapp")  ||
        c.includes("excybe") ||
        c.includes("exclou") ||
        c.includes("exengi") ||
        c.includes("exwrk")  ||
        c.includes("exdata"))   return "AI Certified | Expert";
    if (c.includes("daif"))     return "AI Certified | Fellow";
    if (c.includes("26adv"))    return "AI Cert 2026 Advanced";
    if (c.includes("26aorg"))   return "AI Cert 2026 - Adaptive Orgs Expert";
    if (c.includes("26cyb"))    return "AI Cert 2026 - Cybersecurity Expert";
    if (c.includes("26data"))   return "AI Cert 2026 - Data & Applied AI Expert";
    if (c.includes("26eng"))    return "AI Cert 2026 - Engineering Expert";
    if (c.includes("26fel"))    return "AI Cert 2026 Fellow";
    if (c.includes("26tech"))   return "AI Cert 2026 - Tech Mod";

    return "";
}

// Returns CDN URL for badge image, or empty string if no badge
function getDesignationBadgeUrl(designation) {
    // Replace with your actual CDN URL
    const CDN_BASE = "https://cdn.isg-one.com/Email%20Images";

    if (!designation) return "";

    if (designation === "AI Cert 2026 Advanced")
        return `${CDN_BASE}/AI_Certified_Badge_2026_Advanced_Dark_3X.png`;

    if (designation === "AI Cert 2026 Fellow")
        return `${CDN_BASE}/AI_Certified_Badge_2026_Fellow_Dark_3X.png`;

    if (designation.startsWith("AI Cert 2026 -"))
        return `${CDN_BASE}/AI_Certified_Badge_2026_Expert_Dark_3X.png`;

    return "";
}

// Check if user is Chairman's Club
function isChairmansClub(comments) {
    if (!comments) return false;
    const c = comments.toLowerCase();
    return c.includes("cc") || c.includes("chairman");
}

// Extracts business unit from comments field
// Same logic as C# AddIn: first value before semicolon
function getBusUnit(comments) {
    if (!comments) return "None";
    const parts = comments.split(";");
    const busUnit = parts[0].trim();
    return busUnit || "None";
}

// Extracts country from comments field
// Same logic as C# AddIn: last value after semicolon
function getCountry(comments) {
    if (!comments) return "";
    const parts = comments.split(";");
    if (parts.length > 1) {
        return parts[parts.length - 1].trim();
    }
    return "";
}

// Returns country-specific legal footer HTML
// Matches %%GermanyStart%%...%%GermanyEnd%% and
//         %%SwitzerlandStart%%...%%SwitzerlandEnd%% blocks
function getCountryFooterHtml(country) {
    if (!country) return "";

    const c = country.toLowerCase();

    if (c.includes("germany")) {
        return `
        <p style="line-height:110%;margin:0;">
            <span style="font-size:8pt;color:#767171;">
                Information Services Group Germany GmbH&nbsp;&nbsp;|&nbsp;&nbsp;
                Global Tower, Neue Mainzer Str. 32-36, 60311 Frankfurt am Main
            </span>
        </p>
        <p style="margin:0;">
            <span style="font-size:8pt;color:#767171;">
                Gesch&auml;ftsf&uuml;hrer: Andreas Fahr, Andrea Spiegelhoff&nbsp;&nbsp;|&nbsp;&nbsp;
                Rechtsform: Gesellschaft mit beschr&auml;nkter Haftung
            </span>
        </p>
        <p style="margin:0;">
            <span style="font-size:8pt;color:#767171;">
                Sitz der Gesellschaft: Frankfurt am Main&nbsp;&nbsp;|&nbsp;&nbsp;
                Registergericht und Reg.Nr.: Frankfurt am Main, HRB 94888
            </span>
        </p>
        <p style="margin:4px 0;">
            <span style="font-size:10pt;color:#767171;font-weight:bold;">
                Vertraulichkeitshinweis:
            </span>
            <span style="font-size:8pt;color:#767171;">
                Diese E-Mail enth&auml;lt vertrauliche und rechtlich gesch&uuml;tzte 
                Informationen. Wenn Sie nicht der richtige Adressat sind und diese E-Mail 
                irrt&uuml;mlich erhalten haben, informieren Sie bitte sofort den Absender 
                und vernichten Sie diese E-Mail. Das unerlaubte Ver&ouml;ffentlichen, 
                Kopieren, Weiterleiten oder Nutzen dieser Informationen ist nicht gestattet.
            </span>
        </p>`;
    }

    if (c.includes("switzerland")) {
        return `
        <p style="line-height:110%;margin:0;">
            <span style="font-size:8pt;color:#767171;">
                Information Services Group Switzerland GmbH&nbsp;&nbsp;|&nbsp;&nbsp;
                The Circle 6, 8058 Z&uuml;rich, Switzerland
            </span>
        </p>
        <p style="margin:0;">
            <span style="font-size:8pt;color:#767171;">
                Gesch&auml;ftsf&uuml;hrer: Andreas Fahr, Matth&auml;us K&auml;ppeler&nbsp;&nbsp;|&nbsp;&nbsp;
                Rechtsform: Gesellschaft mit beschr&auml;nkter Haftung
            </span>
        </p>
        <p style="margin:0;">
            <span style="font-size:8pt;color:#767171;">
                Sitz der Gesellschaft: Z&uuml;rich&nbsp;&nbsp;|&nbsp;&nbsp;
                UID: CHE-230.826.460
            </span>
        </p>
        <p style="margin:4px 0;">
            <span style="font-size:10pt;color:#767171;font-weight:bold;">
                Vertraulichkeitshinweis:
            </span>
            <span style="font-size:8pt;color:#767171;">
                Diese E-Mail enth&auml;lt vertrauliche und rechtlich gesch&uuml;tzte 
                Informationen. Wenn Sie nicht der richtige Adressat sind und diese E-Mail 
                irrt&uuml;mlich erhalten haben, informieren Sie bitte sofort den Absender 
                und vernichten Sie diese E-Mail. Das unerlaubte Ver&ouml;ffentlichen, 
                Kopieren, Weiterleiten oder Nutzen dieser Informationen ist nicht gestattet.
            </span>
        </p>`;
    }

    // No country-specific footer for other countries
    return "";
}
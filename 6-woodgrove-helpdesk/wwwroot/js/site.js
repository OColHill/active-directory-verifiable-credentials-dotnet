// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Hill branding: heading, tab title, footer, and removal of demo wording.
document.addEventListener('DOMContentLoaded', function () {
    var h1 = document.querySelector('main h1');
    if (h1 && h1.textContent.indexOf('Welcome to') === 0) {
        h1.textContent = 'Hill IT Helpdesk ID Check';
    }
    document.title = 'Hill IT Helpdesk ID Check';

    var footer = document.querySelector('footer .container');
    if (footer) {
        footer.textContent = '\u00A9 2026 Hill Group IT Department Helpdesk \u2014 ' +
            'No data is retained by this website. Our technicians cannot see your face scan ' +
            'or your data, only the acceptance score. ' +
            'Powered by Microsoft Entra Verified ID Face Check.';
    }

    // Step 3 box: replace the sample's "support personnel" wording (present in DOM, hidden until pass)
    var wrapper = document.getElementById('apps-button-wrapper');
    if (wrapper) {
        var p = wrapper.querySelector('p');
        if (p && p.textContent.indexOf('successfully been verified') !== -1) {
            p.textContent = 'You have been verified. Your IT technician can now see your result and will continue with your request.';
        }
    }

    // The success banner text is written by the app AFTER the check completes - watch and rewrite it
    var msg = document.getElementById('message');
    if (msg) {
        var fix = function () {
            if (msg.textContent.indexOf('if this was a real helpdesk site') !== -1) {
                msg.textContent = 'Verification successful. Your IT technician can now continue with your request.';
            }
        };
        new MutationObserver(fix).observe(msg, { childList: true, subtree: true, characterData: true });
        fix();
    }
});


// Hill: replacement Teams sender - bypasses the sample's tws library, which
// gets blocked by browser CORS preflight against Power Platform flow URLs.
// TO ACTIVATE: paste the flow URL between the quotes below (edit this file in Kudu).
var HILL_TEAMS_WEBHOOK_URL = "";

if (HILL_TEAMS_WEBHOOK_URL) {
    window.sendHelpdeskRequestNotoficationToTeams = function (displayName, email, matchConfidenceScore) {
        var payload = JSON.stringify({
            type: 'message',
            attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: {
                    type: 'AdaptiveCard',
                    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
                    version: '1.4',
                    body: [
                        {
                            type: 'Container',
                            style: 'good',
                            bleed: true,
                            items: [
                                { type: 'TextBlock', text: '\u2714 Face Check passed', weight: 'Bolder', size: 'Medium', color: 'Good' }
                            ]
                        },
                        {
                            type: 'ColumnSet',
                            spacing: 'Medium',
                            columns: [
                                {
                                    type: 'Column',
                                    width: 'auto',
                                    verticalContentAlignment: 'Center',
                                    items: [
                                        { type: 'Image', style: 'Person', width: '40px',
                                          url: 'https://ui-avatars.com/api/?background=0B3041&color=fff&size=64&name=' + encodeURIComponent(displayName) }
                                    ]
                                },
                                {
                                    type: 'Column',
                                    width: 'stretch',
                                    verticalContentAlignment: 'Center',
                                    items: [
                                        { type: 'TextBlock', text: displayName, weight: 'Bolder', size: 'Default', spacing: 'None', wrap: true },
                                        { type: 'TextBlock', text: email, isSubtle: true, spacing: 'None', wrap: true }
                                    ]
                                }
                            ]
                        },
                        {
                            type: 'FactSet',
                            spacing: 'Medium',
                            facts: [
                                { title: 'Match score', value: Math.round(matchConfidenceScore) + '%' },
                                { title: 'Verified', value: new Date().toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) }
                            ]
                        },
                        {
                            type: 'TextBlock',
                            text: '\u21B3 Reply to this card with the ticket number once actioned.',
                            isSubtle: true, size: 'Small', wrap: true, spacing: 'Medium'
                        }
                    ]
                }
            }]
        });
        // Try a proper JSON post first (works if the endpoint accepts the CORS preflight
        // now that the flow allows 'Anyone'); fall back to a no-preflight send if blocked.
        fetch(HILL_TEAMS_WEBHOOK_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: payload
        }).then(function (r) {
            console.log('Hill webhook (json): HTTP ' + r.status);
        }).catch(function (e) {
            console.warn('Hill webhook (json) blocked by browser, retrying no-cors:', e);
            fetch(HILL_TEAMS_WEBHOOK_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: payload
            }).then(function () { console.log('Hill webhook (no-cors): sent, delivery unconfirmed'); });
        });
    };
}

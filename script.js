// ==UserScript==
// @name          Display Unread Badge for OWA
// @namespace     ntrp
// @version       1.0
// @description   This will check for unread messages and put the unread count in the Outlook Web App.
// @match         https://outlook.office.com/*
// @require       https://cdnjs.cloudflare.com/ajax/libs/favico.js/0.3.10/favico.min.js
// ==/UserScript==

var favicon = new Favico({
    animation: 'popFade'
});

function update() {
    var unread = 0;
	try {
	    unread = [...document.querySelectorAll('span[autoId="_n_T"]')]
            .map(node => node.innerHTML)
            .filter(val => val.length > 0)
            .map(val => parseInt(val))
            .reduce((a, b) => a + b, 0);
	} catch(err) {}

    if (unread <= 9) {
        favicon.badge(unread);
    } else {
        favicon.badge("+");
    }
}

window.setInterval(update, 1000);

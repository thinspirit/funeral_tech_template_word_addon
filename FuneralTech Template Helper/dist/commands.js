!function(){var e={};e.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),Office.onReady((function(){})),("undefined"!=typeof self?self:"undefined"!=typeof window?window:void 0!==e.g?e.g:void 0).action=function(e){var n={type:Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,message:"Performed action.",icon:"Icon.80x80",persistent:!0};Office.context.mailbox.item.notificationMessages.replaceAsync("action",n),e.completed()}}();
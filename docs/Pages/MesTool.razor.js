/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
/**
 * Clears the document
 */
export async function clearDocument() {
    await Word.run(async (context) => {
        context.document.body.clear();
    });
}
/**
 * Inserts text at a given location (here hard-coded 😱)
 * @param  {} text
 * @param  {} location
 */
export async function requestContextDemo(text, location) {
    var ctx = new Word.RequestContext();
    var range = ctx.document.getSelection();

    range.insertText(text, "After");

    await ctx.sync();
}

export async function addEvent() {

    let result = "";

    await Word.run(async (context) => {
        console.log("run addEvent");
        const currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        context.document.body.onCommentChanged += (e) => {
            console.log("run onCommentChanged", e);
            
        };

        window.addEventListener("onCommentChanged", (e) => {
            console.log("run onCOntentControllAdded", e);
        })
        /*
        context.document.onSelectionChanged.add((e) => {
            console.log("Event:WordText:", JSON.stringify(body.text));
        });
        */
    })
    .catch(function (error) {
        console.log('Error: ', error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });

    return result;
}

export async function clipboardCopy(text) {
    navigator.clipboard.writeText(text).then(function () {
        alert("クリップボードにコピーしました！");
        return true;
    })
    .catch(function (error) {
        console.error(error);
        return false;
    });
};

export async function getWordText() {

    let result = "";

    await Word.run(async (context) => {
        const currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        var body = context.document.body;
        body.load("text");

        await context.sync();

        result = body.text; //.replaceAll("\r\n","\n").replaceAll("\r","\n");

        //console.log("Paragraph count JS: ");
        console.log("WordText:", JSON.stringify(result));
    })
    .catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });

    return result;
}


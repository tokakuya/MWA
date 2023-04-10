﻿/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
/**
 * Basic function to show how to insert a paragraph at the start of the Word document
 */
export function insertParagraph() {

    return Word.run((context) => {

        // Inserts a paragraph at the start of the document.
        const paragraph = context.document.body.insertParagraph("Hello World from Blazor", Word.InsertLocation.start);

        // Sync the context to run the previous API call, and return.
        return context.sync();
    });
}

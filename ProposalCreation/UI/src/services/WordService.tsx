import { ISection, IDocument } from "../models";
import {IDocumentService } from './IDocumentService';

export class WordService implements IDocumentService
{
    public isDocumentLoaded()
    {
        return Office.context.document.url === null;
    }

    public async updateDocument(document: IDocument)
    {
        
    }

    public async getDocument()
    {
        let result: IDocument;

        return Promise.resolve(result);
    }

    public async getDocuments()
    {
        let result: IDocument;

        return Promise.resolve(result);
    }

    public async getSections()
    {
        return Word.run(async context =>
        {
            let doc = context.document.body.load("paragraphs");
            return context.sync()
            .then(
                () => {
                const styleHeading = "Heading1";
                const styleNormal = "Normal";
                const sections: ISection[] = [];
                //const taskStatus: ITaskStatus = { id: 1, name: "Pending"};
                doc.paragraphs.items.forEach((item, index) =>
                {
                    if (item.text && item.styleBuiltIn === styleHeading) {
                        sections.push({
                            name: item.text,
                            task: "UnAssigned",
                            id: index+"",
                            assignedTo: "",
                            status: 1,
                            content: [],
                            owner: "Foo"
                        });
                    }
                    else if (item.text && item.styleBuiltIn === styleNormal) {
                        sections[sections.length - 1].content.push(item.text);
                    }
                });
        
                return sections;
            });
        });
    }

    public insertText(text: string): void
    {
        Word.run(function (context) {

            let document = context.document;

            document.body.insertText(text, "Start");
     
            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    public getDocumentName()
    {
        let name = Office.context.document.url;
        return name;
    }
}
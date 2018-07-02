import { IDocumentService } from ".";
import { IDocument } from '../models';

export class DocumentService
{
    private wordService: IDocumentService;

    constructor(wordService: IDocumentService)
    {
        this.wordService = wordService;
    }

    public updateDocument(document: IDocument)
    {
        return this.wordService.updateDocument(document);
    }

    public getDocument()
    {
        return this.wordService.getDocument();
    }

    public getDocuments()
    {
        return this.wordService.getDocuments();
    }

    public getSections()
    {
        return this.wordService.getSections();
    }

    public insertText(text: string)
    {
        this.wordService.insertText(text);
    }

    public isDocumentLoaded()
    {
        return this.wordService.isDocumentLoaded();
    }

    public getDocumentName()
    {
        let name = this.wordService.getDocumentName();

        if(name)
        {
            // get file name
            name = name.substr(name.lastIndexOf('/')+1);
            // extract extension
            name = name.substr(0, name.indexOf('.'));
        }

        return name;
    }
}
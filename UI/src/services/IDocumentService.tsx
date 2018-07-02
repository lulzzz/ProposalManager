import {ISection, IDocument} from '../models'

export interface IDocumentService
{
    insertText(text: string);
    getDocumentName();
    getSections(): Promise<ISection[]>;
    isDocumentLoaded(): boolean;
    updateDocument(document: IDocument);
    getDocument(): Promise<IDocument>;
    getDocuments(): Promise<any>;
}
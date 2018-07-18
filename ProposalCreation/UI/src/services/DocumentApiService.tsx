import {ISection, IDocument} from '../models';
import { IDocumentService } from '.';
import { ApiService } from './ApiService';
import * as url from 'url';

export class DocumentApiService implements IDocumentService
{
    private apiService: ApiService;
    private rawDocument: any;
    
    constructor(apiService: ApiService)
    {
        this.apiService = apiService;
    }

    public isDocumentLoaded()
    {
        return Office.context.document.url === null;
    }

    public async updateDocument(document: IDocument)
    {
        // update original full document -- TODO: update IDocument schema to reflect all the properties
        this.rawDocument.proposalDocument.content.proposalSectionList = document.proposalDocument.content.proposalSectionList;
        
        let args = {
            opportunityId: document.id,
            documentData: JSON.stringify(this.rawDocument)
        };
        
        return this.apiService.callApi("Document", "UpdateTask", "POST", args);
    }

    public async getDocument()
    {
        let oppId = this.getOpportunityName();
        let args = { id: oppId};

        return this.apiService.callApi("Document", "GetFormalProposal", "GET", args)
                .then(data => {
                    this.rawDocument = JSON.parse(data.toString());
                    return this.rawDocument as IDocument;
                })
                .catch(error =>
                    {
                        if(error.status == 0)
                        {
                            throw new Error("Connection error.");
                        }
            
                        return error;
                    });
    }

    public async getDocuments()
    {
        let oppId = this.getOpportunityName();
        let args = { id: oppId};

        return this.apiService.callApi("document", "list", "GET", args)
        .then(data => {
            return data;
        })
        .catch(error =>
        {
            if(error.status == 0)
            {
                throw new Error("Connection error.");
            }

            return error;
        });
    }

    public async getSections()
    {
        let oppId = this.getOpportunityName();
        let args = { id: oppId};

        return this.apiService.callApi("Document", "GetFormalProposal", "GET", args)
        .then(data => {
            let doc = JSON.parse(data.toString()) as IDocument
            return doc.proposalDocument.content.proposalSectionList.map(
                item => {
                    let section: ISection;
                    
                    section = {
                        name: item.displayName,
                        task: item.task ? item.task : "Unassigned",
                        id: item.id,
                        assignedTo: item.assignedTo ? item.assignedTo.displayName : "",
                        status: item.sectionStatus,
                        content: [],
                        owner: item.owner.displayName
                    };

                    return section;
                });
            })
            .catch(error =>
                {
                    if(error.status == 0)
                    {
                        throw new Error("Connection error.");
                    }
        
                    return error;
                });
    }

    public insertText(text: string): void
    {
    }

    public getDocumentName()
    {
        return "Foo";//payload.proposalDocument.displayName;
    }

    private getOpportunityName()
    {
        let name = Office.context.document.url;
        let parsedUrl = url.parse(name);
        return parsedUrl.pathname.split('/')[2];
    }
}
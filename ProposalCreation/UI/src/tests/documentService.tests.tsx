import { DocumentService, IDocumentService } from "../services";
import { ISection, IDocument } from "../models";

class WordMockService implements IDocumentService
{
      static DocumentName: string = "TestDocument";
      static Sections: ISection[] = [
            {
                  assignedTo: "Foo",
                  id: "1", 
                  content: [ "hello", "world" ],
                  task: "Content",
                  name: "Section 1",
                  status: 2,
                  owner: "Bar"
            }
      ];

      public getDocument = (): Promise<IDocument> => 
      {
            let result: IDocument;
            return Promise.resolve(result)
      };

      public updateDocument = () => {};

      public insertText(text: string)
      {
            // implement
            return;
      }

      public getDocuments()
      {
            return null;
      }
      
      public getDocumentName()
      {
            const documentUrl = "http://test/";
            return `${documentUrl}/${WordMockService.DocumentName}.docx`;
      }

      public getSections(): Promise<ISection[]>
      {
            return Promise.resolve(WordMockService.Sections);
      }

      public isDocumentLoaded()
      {
            return true;
      }
}

const documentSvc = new DocumentService(new WordMockService());

test("DocumentServiceCanReturnSections", async () => 
{
      const sections = await documentSvc.getSections();
      expect(sections.length).toEqual(1);
      
});

test("DocumentServiceCanReturnDocumentName", () =>
{
      const documentName = documentSvc.getDocumentName();
      expect(WordMockService.DocumentName).toEqual(documentName);
});
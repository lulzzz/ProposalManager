import * as React from 'react';
import { SearchBox, Link } from 'office-ui-fabric-react/lib';
import { ApiService } from '../services/ApiService'
import { DocumentApiService, DocumentService } from '../services'
import {Loading} from './Loading';
import {
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    SelectionMode,
  } from 'office-ui-fabric-react/lib/DetailsList';
import { ScrollablePane } from 'office-ui-fabric-react/lib/ScrollablePane';

export interface IDocument
{
    name: string;
    webUrl: string;
    id: string;
    type: string
}

export interface IDocumentsState
{
    data: IDocument[];
    columns: IColumn[];
    isLoading: boolean
}

export interface IDocumentsProps
{
    token: string
}

export class Documents extends React.Component<IDocumentsProps,IDocumentsState>
{
    private data: IDocument[] = [];
    //private apiService: ApiService;
    private documentService: DocumentService;

    constructor(props: any)
    {
        super(props);

        const columns: IColumn[] = [
            {
                key: 'docIcon',
                name: 'File Type',
                headerClassName: 'Document-header--FileIcon',
                className: 'Document-cell--FileIcon',
                iconClassName: 'Document-Header-FileTypeIcon',
                iconName: 'Page',
                isIconOnly: true,
                fieldName: 'docIcon',
                minWidth: 16,
                maxWidth: 16,
                isPadded: true,
                onRender: this.renderColumn.bind(this)
            },
            {
              key: 'name',
              name: 'Name',
              fieldName: 'name',
              minWidth: 100,
              maxWidth: 200,
              isResizable: false,
              isRowHeader: true,
              isPadded: true,
              isSorted: true,
              isSortedDescending: true,
              onColumnClick: this.onSort.bind(this),
              onRender: this.renderColumn.bind(this)
            },
            {
                key: 'clientDocs',
                name: 'Client Documents',
                fieldName: 'clientDocs',
                minWidth: 100,
                maxWidth: 200,
                isResizable: false,
                onRender: this.renderColumn.bind(this)
            }
        ];

        this.onSort = this.onSort.bind(this);
        this.onChange = this.onChange.bind(this);

        this.state = {
            data: this.data,
            columns: columns,
            isLoading: true
        };

        //this.apiService = new ApiService(this.props.token);
        this.documentService = new DocumentService(new DocumentApiService(new ApiService(this.props.token)));
    }

    componentWillMount()
    {
        this.loadDocuments();
    }

    private renderColumn(item: any, index: number, column: IColumn)
    {
        const fieldContent = item[column.name];

        switch(column.key)
        {
            case 'name':
                return (
                    <Link href={item.webUrl}>{item.name}</Link>
                );
            case 'docIcon':
                return (
                    <img src={ this.getIconImage(item.type) }
                        className = { 'Document-documentIconImage' }
                    />
                );
            case 'clientDocs':
                return (
                    <span>
                      {fieldContent}
                    </span>
                );
            default:
                return fieldContent;
        }
    }

    private getIconImage(type: string) : string
    {
        return `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${type}_16x1.svg`;
    }

    private onSort(e: React.MouseEvent<HTMLElement>, col: IColumn)
    {
        const { columns, data } = this.state;
        const newColumns = columns.slice();
        let sortedData = data.slice();

        if(col.isSortedDescending)
        {
            // Sort the data
            sortedData = sortedData.sort(
                (a: IDocument, b: IDocument) => {
                    if(a.name < b.name)
                    {
                        return 1;
                    }
                    else if(a.name > b.name)
                    {
                        return -1
                    }
                    return 0;
                });
        }
        else
        {
            // Sort the data
            sortedData = sortedData.sort(
                (a: IDocument, b: IDocument) => {
                    if(a.name > b.name)
                    {
                        return 1;
                    }
                    else if(a.name < b.name)
                    {
                        return -1
                    }
                    return 0;
                });
        }
        
        // Update the sort icon
        newColumns[1].isSortedDescending = !col.isSortedDescending;
        this.setState({columns: newColumns, data: sortedData});
    }

    private loadDocuments()
    {
        this.setState({isLoading: true});
        //this.apiService.callApi("document", "list", "GET", null)
        this.documentService.getDocuments()
        .then(
            data => {
                let info: IDocument[];

                info = data as IDocument[]
                
                this.setState({data: info, isLoading: false});
            }
        )
        .catch(err=>{});
    }

    private onChange(text: any): void
    {
        this.setState({ data: text ? this.data.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this.data });
    }

    public render(): JSX.Element
    {
        const { data, columns, isLoading } = this.state;

        if(isLoading)
        {
            return (
                <Loading message="Loading..."/>
            );
        }

        return (
            <ScrollablePane>
                <SearchBox
                    placeholder='Search'
                    onChanged={this.onChange.bind(this)}
                />
                <DetailsList
                    items={ data }
                    columns={ columns }
                    setKey='set'
                    selectionMode={SelectionMode.none}
                    selectionPreservedOnEmptyClick={ true }
                    layoutMode={ DetailsListLayoutMode.justified }
                />
            </ScrollablePane>
        );
    }
}
import * as React from 'react';
import { SearchBox } from 'office-ui-fabric-react/lib';
import { ApiService } from '../services/ApiService'
import {Loading} from './Loading';
import {
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    SelectionMode,
  } from 'office-ui-fabric-react/lib/DetailsList';
import { ScrollablePane } from 'office-ui-fabric-react/lib/ScrollablePane';
import { DocumentService, DocumentApiService } from '../services';

export interface INotesState
{
    notes: INoteIU[];
    isLoading: boolean,
    columns: IColumn[]
}

export interface INotesProps
{
    token: string
}

export interface INoteIU
{
    id: string,
    noteBody: string,
    createdDateTime: Date,
    createdBy: string
}

export class Notes extends React.Component<INotesProps,INotesState>
{
    private documentService: DocumentService;

    constructor(props: any)
    {
        super(props);

        const columns: IColumn[] = [
            {
              key: 'name',
              name: 'Name',
              fieldName: 'noteBody',
              minWidth: 100,
              maxWidth: 200,
              isResizable: false,
              isRowHeader: true,
              isPadded: true,
              isSorted: true,
              isSortedDescending: true,
              onColumnClick: this.onSort.bind(this),
              onRender: this.renderColumn.bind(this)
            }
        ];

        this.onSort = this.onSort.bind(this);
        this.onChange = this.onChange.bind(this);

        this.state = {
            columns: columns,
            isLoading: true,
            notes: []
        };

        this.documentService = new DocumentService(new DocumentApiService(new ApiService(this.props.token)));
    }

    componentWillMount()
    {
        this.loadNotes();
    }

    private renderColumn(item: any, index: number, column: IColumn)
    {
        const fieldContent = item[column.name];
        
        switch(column.key)
        {
            case 'name':
                return (
                    <div>
                        <span className="ms-font-l">{item.noteBody}</span><br/>
                        <span className="ms-font-xs">{this.getFormattedDate(item.createdDateTime)}</span>
                    </div>
                );
            default:
                return fieldContent;
        }
    }

    private getFormattedDate(createdDateTime: Date)
    {
        try{
            const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
            let date = `${createdDateTime.getMonth()}/${createdDateTime.getDate()}/${createdDateTime.getFullYear()} ${dayNames[createdDateTime.getDay()]}`;
            return date;
        }
        catch(e)
        {
            return e;
        }
    }

    private onChange(text: any): void
    {
        const { notes } = this.state;
        this.setState({ notes: text ? notes.filter(i => i.noteBody.toLowerCase().indexOf(text) > -1) : notes });
    }

    private onSort(e: React.MouseEvent<HTMLElement>, col: IColumn)
    {
        const { columns, notes } = this.state;
        const newColumns = columns.slice();
        let sortedData = notes.slice();

        if(col.isSortedDescending)
        {
            // Sort the data
            sortedData = sortedData.sort(
                (a: INoteIU, b: INoteIU) => {
                    if(a.id < b.id)
                    {
                        return 1;
                    }
                    else if(a.id > b.id)
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
                (a: INoteIU, b: INoteIU) => {
                    if(a.id > b.id)
                    {
                        return 1;
                    }
                    else if(a.id < b.id)
                    {
                        return -1
                    }
                    return 0;
                });
        }
        
        // Update the sort icon
        newColumns[1].isSortedDescending = !col.isSortedDescending;
        this.setState({columns: newColumns, notes: sortedData});
    }

    private loadNotes()
    {
        this.setState({isLoading: true});
        
        this.documentService.getDocument()
            .then(document => {
                let notes: INoteIU[];

                notes = document.notes.map(
                    item => {
                        let note: INoteIU;
                        
                        note = {
                            id: item.id,
                            createdBy: item.createdBy.displayName,
                            noteBody: item.noteBody,
                            createdDateTime: new Date(item.createdDateTime)
                        };
    
                        return note;
                    });
                this.setState({ isLoading: false, notes: notes });
            })
            .catch(err => {console.log(err)});
    }

    public render(): JSX.Element
    {
        const { notes, columns, isLoading } = this.state;

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
                    onChanged={this.onChange}
                />
                <DetailsList
                    items={ notes }
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
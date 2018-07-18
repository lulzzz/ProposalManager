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
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DocumentService, DocumentApiService } from '../services';
import { PrimaryButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import './notes.css';
import { ErrorPopup } from './ErrorPopup';
import { LocalizationService } from '../services/LocalizationService';

export interface INotesState
{
    notes: INoteIU[];
    isLoading: boolean;
    columns: IColumn[];
    showPanel: boolean;
    error?: any;
}

export interface INotesProps
{
    token: string;
    localizationService: LocalizationService;
}

export interface INoteIU
{
    id: string,
    noteBody: string,
    createdDateTime: Date,
    createdBy: string;
}

export class Notes extends React.Component<INotesProps,INotesState>
{
    private documentService: DocumentService;
    private currentNote: INoteIU;

    constructor(props: any)
    {
        super(props);
        const { localizationService } = props;

        const columns: IColumn[] = [
            {
              key: 'name',
              name: localizationService.getString("Name"),
              fieldName: 'noteBody',
              minWidth: 100,
              maxWidth: 200,
              isResizable: false,
              isRowHeader: true,
              isPadded: true,
              isSorted: true,
              isSortedDescending: true,
              onColumnClick: this.onSort.bind(this),
              onRender: this.renderColumn.bind(this),
              className: "ms-agile-col-note",
              headerClassName: "ms-agile-col-note"
            },
            {
                key: 'more',
                name: '',
                fieldName: 'more',
                minWidth: 50,
                maxWidth: 50,
                isResizable: false,
                isRowHeader: true,
                isPadded: true,
                onRender: this.renderColumn.bind(this),
                className: "ms-agile-col-btn",
                headerClassName: "ms-agile-col-btn"
            }
        ];

        this.onSort = this.onSort.bind(this);
        this.onChange = this.onChange.bind(this);
        this.closeNotePanel = this.closeNotePanel.bind(this);
        this.activeItemChanged = this.activeItemChanged.bind(this);
        this.openNotePanel = this.openNotePanel.bind(this);

        this.state = {
            columns: columns,
            isLoading: true,
            notes: [],
            showPanel: false
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
        const { localizationService } = this.props;
        
        switch(column.key)
        {
            case 'name':
                return (
                    <div className="ellipsis">
                        <span className="ms-font-l">{item.noteBody}</span><br/>
                        <span className="ms-font-xs">{this.getFormattedDate(item.createdDateTime)}</span>
                    </div>
                );
            case 'more':
                return (
                    <IconButton
                        iconProps={ { iconName: 'MoreVertical' } }
                        title={localizationService.getString("Open")}
                        ariaLabel='MoreVertical'
                        onClick={
                            (e) => {
                                e.preventDefault();
                                this.openNotePanel(item);
                            }
                        }
                    />
                );
            default:
                return fieldContent;
        }
    }

    private openNotePanel(item: any)
    {
        this.setState({ showPanel: true });
    }

    private closeNotePanel()
    {
        this.setState({ showPanel: false });
    }

    private getFormattedDate(createdDateTime: Date)
    {
        try{
            const { localizationService } = this.props;
            
            const dayNames = [localizationService.getString('Sunday'), localizationService.getString('Monday'), localizationService.getString('Tuesday'), localizationService.getString('Wednesday'), localizationService.getString('Thursday'), localizationService.getString('Friday'), localizationService.getString('Saturday')];
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
        newColumns[0].isSortedDescending = !col.isSortedDescending;
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
            .catch(
                err => 
                {
                    this.setState({error: err});
                }
            );
    }

    private activeItemChanged(item?: any, index?: number, ev?: React.FocusEvent<HTMLElement>)
    {
        if(item)
        {
            this.currentNote = item as INoteIU;
        }
    }

    public render(): JSX.Element
    {
        const { notes, columns, isLoading, error } = this.state;
        const { localizationService } = this.props;

        if(error)
        {
            return (
                <ErrorPopup error={error}/>
            );
        }

        if(isLoading)
        {
            return (
                <Loading message={localizationService.getString("Loading")+"..."}/>
            );
        }

        return (
            <div>
                <Panel
                isOpen={ this.state.showPanel }
                type={ PanelType.smallFluid }
                onDismiss={ this.closeNotePanel }
                hasCloseButton={false}
                headerText={localizationService.getString("NoteDetails")}>
                    <div>
                        {/* <span>Created by:</span><br/>
                        <div className="ms-font-m">
                            {this.currentNote ? this.currentNote.createdBy : ""}
                        </div> */}
                        <div className="ms-font-m" style={{paddingTop: "10px"}}>
                            <span>{localizationService.getString("Created")}:</span><br/>
                            {this.currentNote ? this.getFormattedDate(this.currentNote.createdDateTime) : ""}
                        </div>
                        <div className="ms-font-m" style={{paddingTop: "10px"}}>
                            <span>{localizationService.getString("Content")}:</span><br/>
                            {this.currentNote ? this.currentNote.noteBody : ""}
                        </div>
                        <div style={{paddingTop: "10px", display: 'flex'}}>
                            <div style={{paddingLeft: "10px"}}><PrimaryButton onClick={ this.closeNotePanel } text={localizationService.getString("Close")} /></div>
                        </div>
                    </div>
                </Panel>
                <ScrollablePane>
                    <div style={{display: 'flex', paddingTop: '10px'}}>
                        <SearchBox
                            placeholder={localizationService.getString("Search")}
                            onChanged={this.onChange}
                            className="ms-search-notes"
                        />
                        <IconButton
                            iconProps={ { iconName: 'Refresh' } }
                            title={localizationService.getString("Refresh")}
                            ariaLabel='Refresh'
                            onClick={
                                (e) => {
                                    e.preventDefault();
                                    this.loadNotes();
                                }
                            }
                    />
                    </div>
                    <DetailsList
                        items={ notes }
                        columns={ columns }
                        setKey='set'
                        selectionMode={SelectionMode.none}
                        selectionPreservedOnEmptyClick={ true }
                        layoutMode={ DetailsListLayoutMode.justified }
                        onActiveItemChanged={this.activeItemChanged}
                    />
                </ScrollablePane>
            </div>
        );
    }
}
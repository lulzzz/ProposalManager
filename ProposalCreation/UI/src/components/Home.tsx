import * as React from 'react';
import * as $ from 'jquery';
import { DocumentApiService, DocumentService } from '../services'
import {Loading} from './Loading';
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    IColumn,
    SelectionMode,
  } from 'office-ui-fabric-react/lib/DetailsList';
import { ScrollablePane } from 'office-ui-fabric-react/lib/ScrollablePane';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { IDocument, ISection } from '../models';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    ValidationState
  } from 'office-ui-fabric-react/lib/Pickers';
import { assign } from 'office-ui-fabric-react/lib/Utilities';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { ApiService } from '../services/ApiService'

export interface IHomeState
{
    document: IDocument;
    data: ISection[];
    selectionDetails: {};
    columns: IColumn[];
    isLoading: boolean;
    title: string;
    isDocumentLoaded?: boolean;
    peopleList: IPersonaProps[];
    isSaving: boolean;
    showPanel: boolean;
}
import './home.css';

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested Members',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading'
};

export interface IHomeProps
{
    token: string
}

export interface IDialogData
{
    task: string;
    name: string;
}

export class Home extends React.Component<IHomeProps,IHomeState>
{
    private selection: Selection;
    private data: ISection[] = [];
    private documentService: DocumentService;
    private currentSection: ISection;
    private dialogData: IDialogData;
    
    constructor(props: any)
    {
        super(props);
        const columns: IColumn[] = 
        [
            {
                key: 'displayName',
                name: 'Name',
                fieldName: 'displayName',
                minWidth: 100,
                maxWidth: 150,
                isResizable: false,
                isRowHeader: true,
                isPadded: true,
                isSorted: true,
                isSortedDescending: true,
                data: 'string',
                onColumnClick: this.onSort.bind(this),
                onRender: this.renderColumn.bind(this),
                className: "ms-agile-col-name",
                headerClassName: "ms-agile-col-name"
            },
            {
                key: 'task',
                name: 'Task',
                fieldName: 'task',
                minWidth: 100,
                maxWidth: 150,
                isResizable: false,
                isRowHeader: true,
                isPadded: true,
                onRender: this.renderColumn.bind(this),
                className: "ms-agile-col",
                headerClassName: "ms-agile-col"
            },
            {
                key: 'assignedTo',
                name: 'Assigned To',
                fieldName: 'assignedTo',
                minWidth: 100,
                maxWidth: 150,
                isResizable: false,
                isRowHeader: true,
                isPadded: true,
                onRender: this.renderColumn.bind(this),
                className: "ms-agile-col",
                headerClassName: "ms-agile-col"
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
        this.activeItemChanged = this.activeItemChanged.bind(this);
        this.openEditPanel = this.openEditPanel.bind(this);
        this.closePanel = this.closePanel.bind(this);
        this.save = this.save.bind(this);
        
        this.selection = new Selection(
            {
                onSelectionChanged: () => 
                { 
                    this.setState(
                    { 
                        selectionDetails: this.getSelectionDetails()
                    });
                }
          });

        this.state = 
        {
            document: null,
            data: this.data,
            selectionDetails: this.getSelectionDetails(),
            columns: columns,
            isLoading: true,
            title: "Loading",
            peopleList: [],
            isSaving: false,
            showPanel: false
        };

        this.documentService = new DocumentService(new DocumentApiService(new ApiService(this.props.token)));
        this.dialogData = { name: '', task: 'Unassigned'};
    }

    componentDidMount()
    {
        this.loadDocument();
    }

    private renderColumn(item: any, index: number, column: IColumn)
    {
        const fieldContent = item[column.name];

        switch(column.key)
        {
            case 'displayName':
                return (
                    <div className="homeCell ellipsis">
                        <span title={item.name}>{item.name}</span><br/>
                        <span className="ms-font-xs">Status: {item.status}</span><br/>
                        <div className="homeCell ellipsis">
                            <span className="ms-font-xs" title={item.owner}>{item.owner}</span>
                         </div>
                    </div>
                );
            case 'task':
                return (
                    <div className="homeCell">
                        {item.task}
                    </div>
                );
            case 'assignedTo':
                return (
                    <div className="homeCell">
                        {item.assignedTo}
                    </div>
                );
            case 'more':
                return (
                    <IconButton
                        iconProps={ { iconName: 'MoreVertical' } }
                        title='Edit'
                        ariaLabel='MoreVertical'
                        onClick={
                            (e) => {
                                e.preventDefault();
                                this.openEditPanel(item);
                            }
                        }
                    />
                );
            default:
                return <div className="homeCell">{fieldContent}</div>;
        }
    }

    private openEditPanel(item: any)
    {
        this.setState({ showPanel: true });
    }

    private closePanel()
    {
        this.setState({ showPanel: false });
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
                (a: ISection, b: ISection) => {
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
                (a: ISection, b: ISection) => {
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
        newColumns[0].isSortedDescending = !col.isSortedDescending;
        this.setState({columns: newColumns, data: sortedData});
    }

    private getSelectionDetails(): string 
    {
        const selectionCount = this.selection.getSelectedCount();

        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected: ' + (this.selection.getSelection()[0] as any).name;
            default:
                return `${selectionCount} items selected`;
        }
    }
    
    private getTextFromItem(persona: IPersonaProps): string 
    {
        return persona.primaryText as string;
    }

    private validateInput = (input: string): ValidationState => 
    {
        if (input.indexOf('@') !== -1) 
        {
            return ValidationState.valid;
        }
        else if (input.length > 1) 
        {
            return ValidationState.warning;
        } 
        else 
        {
            return ValidationState.invalid;
        }
    }
      
    private async save()
    {
        if(this.currentSection)
        {
            this.setState({isSaving: true});

            const { data, document } = this.state;
            const newData = data.slice();

            let index = newData.findIndex( x => x.id === this.currentSection.id);
            let updatedOpp = Object.assign({}, document);

            if(index > -1)
            {
                newData[index].assignedTo = this.dialogData.name;
                newData[index].task = this.dialogData.task;
                this.currentSection = newData[index];

                //update document object
                let sectionToUpdate = updatedOpp.proposalDocument.content.proposalSectionList.find(x => x.id == this.currentSection.id);

                sectionToUpdate.task = this.dialogData.task;
                sectionToUpdate.assignedTo = updatedOpp.teamMembers.find(x => x.displayName === this.dialogData.name);

                await this.documentService.updateDocument(updatedOpp);
            }

            $("#popup").removeClass("on").addClass("off");
            this.setState({isSaving: false, data: newData, showPanel: false, document: updatedOpp});
        }
    }

    private onTaskChange(e: React.ChangeEvent<HTMLSelectElement>)
    {
        assign(this.dialogData, { task: e.target.value });
    }

    private onPeoplePickerChange(items: IPersonaProps[])
    {
        let name = "";
        // A Persona has been selected
        if(items && items.length > 0)
        {
            name = items[0].primaryText;
        }
        
        assign(this.dialogData, { 
            name: name
        });
    }

    private getDefaultSelected(assignedTo:string)
    {
        let { peopleList } = this.state;

        const result = peopleList.find(x =>x.primaryText === assignedTo)

        if(result)
        {
            return [result as IPersonaProps];
        }
        else
        {
            return [];
        }
    }

    private renderPeoplePicker(assignedTo: string) 
    {
        return (
            <div className="homeCell">
                <CompactPeoplePicker
                    itemLimit={1}
                    defaultSelectedItems={this.getDefaultSelected(assignedTo)}
                    onResolveSuggestions={ this.onFilterChanged }
                    getTextFromItem={ this.getTextFromItem }
                    className={ 'ms-PeoplePicker' }
                    pickerSuggestionsProps={ suggestionProps }
                    onValidateInput={ this.validateInput }
                    onChange={ this.onPeoplePickerChange.bind(this)}
                />
            </div>
        );
    }

    private listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) 
    {
        if (!personas || !personas.length || personas.length === 0) 
        {
            return false;
        }
        
        return personas.filter(item => item.primaryText === persona.primaryText).length > 0;
    }

    private removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) 
    {
        return personas.filter(persona => !this.listContainsPersona(persona, possibleDupes));
    }

    private doesTextStartWith(text: string, filterText: string): boolean 
    {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    }

    private filterPersonasByText(filterText: string): IPersonaProps[] 
    {
        let { peopleList } = this.state;
        return peopleList.filter(item => this.doesTextStartWith(item.primaryText as string, filterText));
    }

    private onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => 
    {
        if (filterText) 
        {
            let filteredPersonas: IPersonaProps[] = this.filterPersonasByText(filterText);

            filteredPersonas = this.removeDuplicates(filteredPersonas, currentPersonas);
            return limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
        } 
        else 
        {
            return [];
        }
    }

    private loadDocument()
    {
        this.setState({isLoading: true});
        
        this.documentService.getDocument()
            .then(document => {
                const documentName = document.displayName;

                let sections: ISection[];

                sections = document.proposalDocument.content.proposalSectionList.map(
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

                const peopleList: IPersonaProps[] = [];
                document.teamMembers
                    .forEach(
                        (item, index) => 
                        {
                            let names = item.displayName.split(' ');
                            let initials = names.length > 1 ? `${names[0].charAt(0)}${names[1].charAt(0)}` : names[0].charAt(0);
                            let persona = {
                                    key: index,
                                    imageInitials: initials,
                                    primaryText: item.displayName
                                }
                            
                            peopleList.push(persona);
                        }
                    );
            
                this.setState({document: document, data: sections, isLoading: false, title: documentName, isDocumentLoaded: true, peopleList: peopleList });
            })
            .catch(err => {console.log(err)});
    }

    private onChange(text: any): void
    {
        this.setState({ data: text ? this.data.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this.data });
    }

    private activeItemChanged(item?: any, index?: number, ev?: React.FocusEvent<HTMLElement>)
    {
        if(item)
        {
            this.currentSection = item as ISection;
            assign(this.dialogData, { name: this.currentSection.assignedTo, task: this.currentSection.task });
        }
    }

    public render(): JSX.Element
    {
        const { data, columns, isLoading, title, isDocumentLoaded, isSaving } = this.state;

        if(isLoading)
        {
            return (
                <Loading message="Loading..."/>
            );
        }

        if(isSaving)
        {
            return (
                <Loading message="Saving..." overlay={true}/>
            );
        }

        if(isDocumentLoaded && isDocumentLoaded.valueOf() === false)
        {
            return <div>You must open a document.</div>;
        }

        return (
            <div>
                <Panel
                isOpen={ this.state.showPanel }
                type={ PanelType.smallFluid }
                // tslint:disable-next-line:jsx-no-lambda
                onDismiss={ this.closePanel }
                hasCloseButton={false}
                headerText={this.currentSection ? this.currentSection.name : ""}>
                    <div>
                        <div className="ms-font-m">
                            <span>Owner:</span><br/>
                            {this.currentSection ? this.currentSection.owner : ""}
                        </div>
                        <div className="ms-font-m" style={{paddingTop: "10px"}}>
                            <span>Task:</span><br/>
                            <select style={{height: "32px"}} 
                                defaultValue={this.currentSection ? this.currentSection.task : "Unassigned"}
                                onChange={this.onTaskChange.bind(this)}
                            >
                                <option value="Approval">Approval</option>
                                <option value="Content">Content</option>
                                <option value="Unassigned">Unassigned</option>
                            </select>
                        </div>
                        <div className="ms-font-m" style={{paddingTop: "10px"}}>
                            <span>Assigned To:</span><br/>
                            {this.renderPeoplePicker(this.currentSection ? this.currentSection.assignedTo : "")}
                        </div>
                        <div style={{paddingTop: "10px", display: 'flex'}}>
                            <div><PrimaryButton onClick={ this.save } text='Update' /></div>
                            <div style={{paddingLeft: "10px"}}><DefaultButton onClick={ this.closePanel } text='Cancel' /></div>
                        </div>
                    </div>
                </Panel>
                <ScrollablePane>
                    <h1 className='ms-font-xxl'>{title}</h1>
                    <MarqueeSelection selection={ this.selection }>
                        <DetailsList
                            items={ data }
                            columns={ columns }
                            setKey='set'
                            selection={ this.selection }
                            selectionMode={SelectionMode.none}
                            selectionPreservedOnEmptyClick={ true }
                            layoutMode={ DetailsListLayoutMode.justified }
                            onActiveItemChanged={this.activeItemChanged}
                            className="homeCell"
                        />
                    </MarqueeSelection>
                </ScrollablePane>
        </div>
        );
    }
}
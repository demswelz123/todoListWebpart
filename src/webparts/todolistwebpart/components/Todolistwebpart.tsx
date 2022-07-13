import * as React from 'react';
import styles from './Todolistwebpart.module.scss';
import { ITodolistwebpartProps } from './ITodolistwebpartProps';
import { cloneDeep, escape, update } from '@microsoft/sp-lodash-subset';
import { DayOfWeek, PrimaryButton, List, DefaultButton, getIconClassName, DialogType, ActivityItem, Dialog, Panel, TextField, Dropdown, IDropdownOption, DatePicker, PanelType, Spinner, SpinnerSize, Pivot, PivotItem, PivotLinkFormat, Checkbox } from 'office-ui-fabric-react';
import { Items } from '@pnp/sp';

import { sp } from '@pnp/sp';
import { isArray } from '@pnp/common';
import ErrorHandlingField from './common/ErrorHandlingField';

export interface ITodolistwebpartState {
    isProcessing: boolean;
    items: any[];
    showPanel: boolean;
    showModal: boolean;
    activeItem: any;
    tempItem: any;
    activeIndex: number;
    errorMsg: any;
    saveReady: boolean;
    subTask: any[];
    editFlag: boolean;
    tempSubtask: any;



}


const REQUIRED = [

    "Title",
    "Status",
    "DueDate"
];
const LOREM = [
    {
        Name: '1 kilo mais',
        Status: false

    },
    {
        Name: '1 kilo oil',
        Status: false

    },
    {
        Name: '1 kilo salt',
        Status: false

    }
];
export default class Todolistwebpart extends React.Component<ITodolistwebpartProps, ITodolistwebpartState> {

    constructor(props) {
        super(props);

        this.state = {
            isProcessing: false,
            showModal: false,
            showPanel: false,

            tempItem: {
                Title: '',
                Description: '',
                Status: 'Not Started',
                DueDate: new Date(),
            },
            tempSubtask: {
                Title: '',
                Status: 'Not Started',
                DateCompleted: null,
                SubtestID: null
            },

            items: [],
            activeItem: null,
            activeIndex: -1,
            errorMsg: {},
            saveReady: false,
            subTask: [],
            editFlag: false,

        };
    }

    private _checkIsFormReady = () => {
        let { errorMsg, tempItem } = this.state;
        REQUIRED.forEach(field => {



            if (!tempItem[field] || (typeof tempItem[field] == "string" && tempItem[field].trim() === '') ||
                (isArray(tempItem[field]) && tempItem[field].lenght == 0)) {
                errorMsg[field] = errorMsg[field] || 'this field must not be empty';

            } else {
                errorMsg[field] = null;
            }


        });

        let flag = true;
        for (let k of Object.keys(errorMsg)) {

            if (errorMsg[k]) {
                flag = false;
                break;
            }
        }
        this.setState({ errorMsg, saveReady: flag });

    }


    public componentDidMount(): void {
        sp.web.lists.getById('701bcceb-c127-4065-8607-390687788696').items.get()
            .then(res => {

                const items = [];

                res.forEach(item => {
                    const temp = {
                        ID: item.ID,
                        Title: item.Title,
                        Description: item.Description,
                        Status: item.Status || 'Not Started',
                        DueDate: item.DueDate || new Date()
                    };

                    items.push(temp);
                });
                this.setState({ items });
            });

    }
    public render(): React.ReactElement<ITodolistwebpartProps> {
        const { items, showModal, subTask, tempSubtask, activeItem, activeIndex, tempItem, showPanel, isProcessing, saveReady, errorMsg } = this.state;

        return (
            <div className="ms=Grid ">

                <div className="ms-Grid -row">


                    <div className={"ms-Grid-col ms-sm12 " + styles.centerMass}>
                        <b>  <span > TODO LIST </span></b>
                        <br /> <br />
                    </div>

                    <div className="ms-Grid-col ms-sm12" >
                        <PrimaryButton
                            text="ADD ITEM"
                            onClick={() => {
                                // static
                                // const item={
                                //     Task: ' test',
                                //     Description: ' "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."',
                                //     Status:' not started',
                                //     DueDate: new Date().toLocaleString()

                                // };   

                                // items.push(item);

                                // this.setState({ items });

                                this.setState({ showPanel: true });

                            }}
                        />
                        <br /> <br />

                    </div>
                    <div className="ms-Grid-col ms-sm12" >
                        <List
                            items={cloneDeep(items)}
                            onRenderCell={(item?: any, index?: number, isScrolling?: boolean) => {

                                return (
                                    <div className="ms-Grid-col ms-sm12" style={{ marginBottom: '10px', border: '1px ridge blue' }}>

                                        <div className="ms-Grid-col ms-sm8" >

                                            <div className="ms-Grid-col ms-sm12" >
                                                ID:{item.ID}
                                            </div>
                                            <div className="ms-Grid-col ms-sm12">
                                                Title:{item.Title}
                                            </div>
                                            <div className="ms-Grid-col ms-sm12">
                                                Status:{item.Status}
                                            </div>
                                            <div className="ms-Grid-col ms-sm12">
                                                Due Date: {item.DueDate.toLocaleString()}
                                            </div>
                                        </div>
                                        <div className="ms-Grid-col ms-sm4">
                                            <div className="ms-Grid-col ms-sm12" style={{ margin: '5px auto' }}>
                                                <DefaultButton
                                                    style={{ background: '#00b7c3' }}
                                                    iconProps={{ iconName: 'View' }}
                                                    onClick={() => {

                                                        item.DueDate = new Date(item.DueDate);

                                                        this.setState({
                                                            tempItem: item,
                                                            showPanel: true,
                                                            editFlag: true
                                                            // showModal:true,
                                                            // activeItem: item,
                                                            // activeIndex: index
                                                        });


                                                    }}


                                                />
                                            </div>
                                            <br /><br />
                                            <div className="ms-Grid-col ms-sm12" style={{ margin: '5px auto' }}>
                                                <DefaultButton
                                                    style={{ background: '#d83b01' }}
                                                    iconProps={{ iconName: 'Delete' }}
                                                    onClick={() => {

                                                        this.setState({ isProcessing: true, });

                                                        sp.web.lists.getById('701bcceb-c127-4065-8607-390687788696').items.getById(item.ID)
                                                            .recycle().then(_ => {

                                                            });
                                                        const res = items.filter((it, num) => {
                                                            if (index != num) {
                                                                return it;
                                                            }

                                                        });

                                                        this.setState({ items: cloneDeep(res), isProcessing: false });

                                                    }}
                                                    disabled={isProcessing}
                                                />

                                            </div>



                                        </div>



                                    </div>



                                );

                            }}
                        />
                    </div>


                </div>

                <Panel

                    isOpen={showPanel}
                    onOuterClick={() => { }}
                    type={PanelType.medium}
                    onDismiss={() => this.setState({ showPanel: false })}
                >
                    {this._handleRenderHeader()}

                    <Pivot linkFormat={PivotLinkFormat.links}>
                        <PivotItem headerText="Task details">
                            <div className="ms-Grid-col ms-sm12" style={{ margin: '10px 0' }}>

                                <ErrorHandlingField

                                    isRequired={true}
                                    label="Title"
                                    errorMessage={errorMsg.Title}
                                    parentClass={"ms-Grid-col ms-sm12"}

                                >
                                    <TextField
                                        value={tempItem.Title}
                                        onChanged={(newVal: string) => {
                                            tempItem.Title = newVal;

                                            this.setState({ tempItem }, () => {
                                                this._checkIsFormReady();
                                            });

                                        }}
                                    />

                                </ErrorHandlingField>

                                <ErrorHandlingField
                                    isRequired={false}
                                    label="Description"
                                    errorMessage={errorMsg.Description}
                                >
                                    <TextField

                                        value={tempItem.Description}
                                        onChanged={(newVal: string) => {
                                            tempItem.Description = newVal;
                                            this.setState({ tempItem }, () => {
                                                this._checkIsFormReady();
                                            });
                                        }}
                                        multiline
                                        rows={6}
                                    />

                                </ErrorHandlingField>

                                <ErrorHandlingField
                                    isRequired={true}
                                    label="Status"
                                    errorMessage={errorMsg.Status}
                                >
                                    <Dropdown
                                        options={[

                                            { key: 'Not Started', text: 'Not Started' },
                                            { key: 'In-Progress', text: 'In Progress' },
                                            { key: 'On-Hold', text: 'On-Hold' },
                                            { key: 'Completed-', text: 'Completed' },

                                        ]}
                                        selectedKey={tempItem.Status}
                                        onChanged={(option: IDropdownOption, index?: number) => {
                                            tempItem.Status = option.key;

                                            this.setState({ tempItem }, () => {
                                                this._checkIsFormReady();
                                            });

                                        }}
                                    />
                                </ErrorHandlingField>

                                <ErrorHandlingField
                                    isRequired={true}
                                    label="Due Date"
                                    errorMessage={errorMsg.DueDate}
                                >
                                    <DatePicker
                                        value={tempItem.DueDate}
                                        onSelectDate={(date: Date) => {
                                            this.setState({ tempItem }, () => {
                                                this._checkIsFormReady();
                                            });

                                        }}
                                    />

                                </ErrorHandlingField>
                            </div>
                        </PivotItem>
                        <br />

                        <PivotItem headerText="Subtasks">

                            <div className="ms-Grid-col ms-sm12" style={{ margin: '10px 0' }}>
                                <div>
                                    <PrimaryButton
                                        text="add Sub-Task"
                                        onClick={() => {

                                            this.setState({
                                                showModal: true,

                                            });
                                        }}
                                    />
                                    <br /><br />


                                </div>
                            </div>
                            <div className="ms-Grid-col ms-sm12" >
                                <List
                                    items={cloneDeep(this.state.subTask)}
                                    onRenderCell={(item?: any, index?: number, isScrolling?: boolean) => {
                                        const d = new Date().toLocaleDateString();

                                        return (
                                            <div className="ms-Grid-col ms-sm12" style={{ marginBottom: '10px', border: '1px ridge blue' }}>

                                                <div className="ms-Grid-col ms-sm8" >

                                                    <div className="ms-Grid-col ms-sm12" style={item.Status ? { textDecoration: 'line-through' } : {}}>
                                                        {item.Title}
                                                    </div>
                                                    <div className="ms-Grid-col ms-sm12 ">
                                                        Status: {item.Status != "Not Started" ? "Done" : "Pending"}
                                                    </div>
                                                    <div className="ms-Grid-col ms-sm12 ">
                                                        Date Completed: {item.Status != "Not Started" ? d : "N/A"}
                                                    </div>
                                                </div>
                                                <div className="ms-Grid-col ms-sm4">
                                                    <div className="ms-Grid-col ms-sm12" style={{ margin: '5px auto' }}>
                                                        <div className="ms-Grid-col ms-sm2">
                                                            <Checkbox
                                                                label="Mark as done"
                                                                style={{ background: '#00b7c3' }}
                                                                onChange={(ev, checked: boolean) => {
                                                                    const temp = this.state.subTask;
                                                                    temp[index].Status = checked;

                                                                    this.setState({ subTask: temp });

                                                                }}
                                                                value={item.Status}
                                                            />
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        );
                                    }}
                                />
                            </div>


                            {/* <div className="ms-Grid-col ms-sm12" >
                                <List

                                    items={cloneDeep(this.state.subTask)}

                                    onRenderCell={(item?: any, index?: number, isScrolling?: boolean) => {

                                        const d = new Date().toLocaleDateString();
                                        return (
                                            <div
                                                className="ms-Grid-col ms-sm12"
                                                style={{ marginBottom: "10px", border: "1px ridge black" }}
                                            >
                                                <div className="ms-Grid-col ms-sm8">
                                                    <div className="ms-Grid-col ms-sm12 ">
                                                        Name: {item.Title}
                                                    </div>

                                                </div>
                                                <div className="ms-Grid-col ms-sm12 ">
                                                    Status: {item.Status != "Not Started" ? "Done" : "Pending"}
                                                </div>
                                                <div className="ms-Grid-col ms-sm12 ">
                                                    Date Completed: {item.Status != "Not Started" ? d : "N/A"}
                                                </div>
                                            </div>
                                        );
                                    }}
                                />
                            </div> */}
                            {/* <div className="ms-Grid-col ms-sm12">
                                    <div className="ms-Grid-col ms-sm12" style={{ margin: " 5px auto" }}>
                                        <div className="ms-Grid-col ms-sm12">
                                            <Checkbox
                                                label="Mark as done"
                                                style={{ background: 'red', width: '100%', padding: '15px' }}
                                                onChange={(ev, checked: boolean) => {
                                                    const temp = this.state.subTask;
                                                    temp[index].Status = checked;

                                                    this.setState({ subTask: temp });
                                                }}
                                                value={item.Status}
                                            />
                                        </div>
                                    </div>
                                </div>
                            </div> */}




                        </PivotItem>

                    </Pivot>






                    {this._handleRenderFooter()}
                </Panel>

                <Dialog
                    hidden={!showModal}
                    modalProps={{ isBlocking: true }}
                    onDismiss={() => this.setState({
                        showModal: false,
                        tempSubtask: {
                            Title: '',
                            Status: 'Not Started',
                            DateCompleted: null,
                            SubtestID: null
                        }, activeIndex: -1
                    })
                    }
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Add Sub-Task',
                    }}
                >
                    <div className="ms-Grid-col ms-sm12">
                        <ErrorHandlingField

                            isRequired={true}
                            label="Title"
                            errorMessage={errorMsg.Title}
                            parentClass={"ms-Grid-col ms-sm12"}

                        >
                            <TextField
                                value={tempSubtask.Title}
                                onChanged={(newVal: string) => {
                                    tempSubtask.Title = newVal;

                                    this.setState({ tempSubtask }, () => {

                                    });

                                }}
                            />
                        </ErrorHandlingField>
                    </div>
                    <br />
                    <div className="ms-Grid-col ms-sm12">
                        <div className="ms-Grid-col ms-sm4">
                            <PrimaryButton
                                text="Save"
                                style={{ width: "100%" }}
                                onClick={() => {
                                    subTask.push(tempSubtask);

                                    this.setState({ showModal: false, subTask }, () => {
                                        this._checkIsFormReady();
                                    });
                                }}
                            />


                        </div>
                        <br />
                        <div className="ms-Grid-col ms-sm4">
                            <DefaultButton
                                text="Cancel"
                                style={{ width: "100%" }}
                                onClick={() => {
                                    this.setState({
                                        showModal: false,
                                        tempSubtask: {
                                            Title: '',
                                            Status: 'Not Started',
                                            DateCompleted: null,
                                            SubtestID: null
                                        }
                                    });
                                }}
                            />


                        </div>
                    </div>
                    {/* <div className='ms-Grid-col ms-sm12'>
                <span style={{ textAlign:'center' }}>
                    {activeItem && (
                        <div>
                        <b>Description:</b> {activeItem.Description}
                        </div>
                    )}
                </span>
            </div> */}


                </Dialog>

            </div>
        );
    }

    private _handleRenderHeader = () => {
        return (
            <div className={styles.siteTheme + " ms-Grid-row " + styles.panelHeaderV2} style={{ display: 'flex' }}>
                <div className={"ms-Grid-col ms-sm12 " + styles.awkwardSmtoMdHeader}>
                    <div> NEW TODO FORM </div>
                </div>
                {this.state.tempItem.Status && (

                    <div className={"ms-Grid-col ms-sm12 ms-xl6 " + styles.awkwardSmtoMdStatus}>
                        <div>{`status: ${this.state.tempItem.Status} `}</div>

                    </div>

                )}
            </div>

        );


    }

    private _handleRenderFooter = () => {
        const { tempItem, items, saveReady, isProcessing, editFlag } = this.state;
        return (
            <div className="ms-Grid-row"  >
                <div className="ms-Grid-row" >
                    <div className={"ms-Grid-col ms-sm6 ms-xl6 " + styles.awkwardMdtoLg3} >
                        <PrimaryButton
                            text="Save"
                            style={{ width: '50%', marginTop: '3%' }}
                            onClick={async () => {
                                this.setState({ isProcessing: true });

                                if (editFlag) {

                                    await sp.web.lists.getById('701bcceb-c127-4065-8607-390687788696').items.getById(tempItem.ID)
                                        .update(tempItem).then(res => {
                                            const temp = items.map((i, n) => {
                                                if (i.ID == tempItem.ID) {
                                                    return tempItem;

                                                } else {
                                                    return i;

                                                }

                                            });
                                            this.setState({
                                                items: temp, showPanel: false, editFlag: false, isProcessing: false,
                                                tempItem: {
                                                    Title: '',
                                                    Description: '',
                                                    Status: 'Not Started',
                                                    DueDate: new Date()
                                                }
                                            });
                                        });

                                } else {
                                    await sp.web.lists.getById('701bcceb-c127-4065-8607-390687788696').items.add(tempItem).then(res => {
                                        items.push(tempItem);
                                        this.setState({
                                            items, showPanel: false, editFlag: false, isProcessing: false,
                                            tempItem: {
                                                Title: '',
                                                Description: '',
                                                Status: 'Not Started',
                                                DueDate: new Date()
                                            }
                                        });
                                    });
                                }


                            }}
                            disabled={!saveReady || isProcessing}

                        >

                            {isProcessing && (
                                <Spinner
                                    size={SpinnerSize.small}
                                    style={{ marginLeft: "5px" }}
                                />
                            )}
                        </PrimaryButton>
                    </div>
                </div>

                <div className={"ms-Grid-col ms-sm6 ms-xl3 " + styles.awkwardMdtoLg3} >
                    <DefaultButton
                        style={{ width: '40%', marginLeft: '55%', marginTop: '-32px' }}

                        text="Cancel"
                        onClick={() => {
                            this.setState({
                                showPanel: false, editFlag: false, tempItem: {
                                    tempItem: {
                                        Title: '',
                                        Description: '',
                                        Status: 'Not Started',
                                        DueDate: new Date()
                                    }
                                }
                            });
                        }}
                        disabled={isProcessing}
                    />
                </div>
            </div>


        );
    }
}

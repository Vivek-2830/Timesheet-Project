import * as React from 'react';
import styles from './TimeSheet.module.scss';
import { ITimeSheetProps } from './ITimeSheetProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, DialogType, Dropdown, Icon, IIconProps, PrimaryButton, TextField } from 'office-ui-fabric-react';
import * as moment from 'moment';
import {  PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 

export interface ITimeSheetState {
  Title: any;
  ClientName: any;
  Status: any;
  Technology: any;
  AssignTo: any;  
  Modified: any;
  AllProjects: any;
  AllChoice: any;
  AddFilterDialog: boolean;
  AssignToID: any;
  ProjectStatuslist: any;
  Technologylist: any;
  Client: any;
  EditProjectTitle: any;
  EditProjectClientName: any;
  EditProjectStatus: any;
  EditProjectTechnology: any;
  EditProjectAssignTo: any;
  EditProjectModified: any;
  EditProjectAssignToID: any;
  EditFilterDialog: boolean;
  CurrenttimeSheetID : any;
  ClientID: any;
  EditProjectClientNameId: any;
  DeleteFilterDialog: boolean;
  DeleteCurrentitem: any; 
}

require("../assets/css/style.css");
require("../assets/css/fabric.min.css");

const dialogContentProps = {
  title: "Add Employee Details.!!",
};

const EditFilterDialogContentProps = {
  title: "Update News",
};

const DeleteFilterDialogContentProps = {
  // type: DialogType.normal,
  // title: "Alert",
  // // closeButtonAriaLabel: 'Close',
  // subText: "Are You Sure Delete this Details.?"
};

const addmodelProps = {
  className: "Add-Dialog"
};

const updatemodelProps = {
  className : "Update-Dialog"
};

const deletmodelProps = {
  className : "Delete-Form"
};

const TextDocumentEdit : IIconProps = { iconName: 'TextDocumentEdit' };

const addIcon: IIconProps = { iconName: 'Add' };
const SendIcon : IIconProps = { iconName: 'Send'};
const CancelIcon : IIconProps = { iconName: 'Cancel'};
const EventDate: IIconProps = { iconName: 'EventDate' };
const People: IIconProps = { iconName: 'People' };

export default class TimeSheet extends React.Component<
ITimeSheetProps, 
ITimeSheetState
> {
  constructor(props: ITimeSheetProps, state: ITimeSheetState) {
    super(props);

    this.state = {
      Title: "",
      ClientName: "",
      Status: "",
      Technology: "",
      AssignTo: "",
      Modified: "",
      AllProjects: "",
      AddFilterDialog: true,
      AssignToID: [],
      AllChoice: "",
      ProjectStatuslist: "",
      Technologylist: "",
      Client :"",
      EditProjectTitle: "",
      EditProjectClientName: "",
      EditProjectStatus: "",
      EditProjectTechnology: "",
      EditProjectAssignTo: "",
      EditProjectModified: "",
      EditFilterDialog: true,
      CurrenttimeSheetID :"",
      ClientID: "",
      EditProjectAssignToID: "",
      EditProjectClientNameId: "",
      DeleteFilterDialog: true,
      DeleteCurrentitem: ""
    };
  }
  
  public render(): React.ReactElement<ITimeSheetProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
        <div className="timeSheet">
          <div className='ms-Grid'>
            <div className='ms-Grid-row'>
              <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg10 ms-xl12">
                <div className='d-flex-header'>
                  <h3 className='Time-Title'>Employee Details</h3>
                  <div className='Add-Project'>
                    <PrimaryButton className='Add-sheet' iconProps={addIcon} type='Add' text='Add'  onClick={() => this.setState({ AddFilterDialog: false })} />
                  </div>
                </div>
              </div>
              
              <div className='ms-Grid-row '>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  {
                  this.state.AllProjects.length > 0 &&  
                    this.state.AllProjects.map((item) => {
                    return (
                      <div className='ms-Grid-col ms-sm6 ms-md3 ms-lg3'>
                        <div className='Card'>
                          
                          <div className='card-AssignTo'>
                            <h3>{ item.AssignTo ? item.AssignTo.Title: "" }</h3>
                          </div>  
                          <div className='E-Icon'>
                            <Icon className="Edit-Icon" iconName='Edit' onClick={() => this.setState({ EditFilterDialog : false, CurrenttimeSheetID : item.ID, ClientID : item.ClientName.ID },() => this.GetEditProject(item.ID))} />
                          </div>
                          <div className="D-Icon">
                            <Icon className='Delete-Icon' iconName='Delete' onClick={() => this.setState({ DeleteFilterDialog : false, DeleteCurrentitem: item.ID })}/>
                          </div>
                          <div className='modified'>
                          <p><Icon className='Date-Icon' iconName="EventDate"></Icon>{moment(new Date(item.Modified)).format("DD MMM,YY")}</p>
                          </div>
                          <div className='Card-Title'>
                              <span><Icon className="People-Icon" iconName="People"/>{item.Title}</span>
                          </div>
                          <div>
                              <p><Icon className="People-Icon" iconName='People'/>{ item.ClientName ? item.ClientName.Title: "" }</p>
                          </div>
                          
                          <div className='card-Choice'>
                            <p><Icon iconName='Devices3' className='Device-icon'></Icon>{item.Status}</p>
                            <p><Icon iconName='Devices3' className='Device-icon'></Icon>{item.Technology}</p>
                          </div>

                          <div>
                          </div>
                        </div>
                      </div>
                    );
                    })
                  }
                </div>
              </div>

              <Dialog
                    hidden={this.state.AddFilterDialog}
                    onDismiss={() =>
                      this.setState({
                        AddFilterDialog : true,
                        Title: "",
                        ClientName : "",
                        AssignTo: "",
                        Status: "",
                        Technology: ""
                      })}
                    dialogContentProps={dialogContentProps}
                    modalProps={addmodelProps}
                    minWidth={500}
              >
            
              <div>
                <div className='ms-Grid-row ms-md'>
                  <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>

                    <div className='Add-Title'>
                      <TextField
                        label="Title"
                        name="Title"
                        type="Text"
                        required={true}
                        onChange={(value) => 
                          this.setState({Title: value.target["value"]})
                        }
                        value={this.state.Title}
                      />
                    </div>
                    
                    <div className='Add-ClientName'>
                      <TextField
                        label="ClientName"
                        name="ClientName"
                        type="Text"
                        placeholder="Please Enter Your Name"
                        required={true}
                        onChange={(value) =>
                          this.setState({ClientName: value.target["value"]})
                        }
                        value={this.state.ClientName.Title}
                      />
                    </div>

                    <div className='Add-ProjectList'>
                      <Dropdown
                        options={(this.state.ProjectStatuslist)}
                        label="Status"
                        placeholder="Select Status"
                        required
                        onChange={(e, option, index) =>
                          this.setState({ Status: option.text})
                        }
                      />
                    </div>

                    <div className='Add-Technology'>
                      <Dropdown 
                        options={(this.state.Technologylist)}
                        label="Selct Technology"
                        placeholder="Select Technology"
                        required
                        onChange={(e, option, index) => 
                          this.setState({ Technology: option.text})
                        }                      
                      />
                    </div>
                  
                    <div className='Add-Context'>
                      <PeoplePicker 
                          context={this.props.context}
                          titleText="People Picker"
                          personSelectionLimit={3}
                          // groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
                          showtooltip={true}
                          required={true}
                          defaultSelectedUsers={this.state.AssignTo.Title}
                          onChange={this._getPeoplePickerItems}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={300} 
                          ensureUser={true}
                      />
                    </div>

                  </div>
                </div>
                <div className='ms-Grid-row Add-Details'>
                    <div className='ms-Grid-col Submit-Details'>
                        <PrimaryButton
                            iconProps={SendIcon}
                            type="Submit"
                            text="Submit"
                            onClick={() => this.AddProject()}
                        />
                    </div>
                    <div className="ms-Grid-col Cancel-Details">
                      <DefaultButton
                            iconProps={CancelIcon}  
                            type="Cancel"
                            text="Cancel"
                            onClick={() => this.setState({ AddFilterDialog : true })}
                      />
                    </div>
                </div>
              </div>
              </Dialog>
                  
              <Dialog
                    hidden={this.state.EditFilterDialog}
                    onDismiss={() =>
                      this.setState({
                        EditFilterDialog : true,
                        Title: "",
                        ClientName : "",
                        AssignTo: "",
                        Status: "",
                        Technology: ""
                      })}
                    dialogContentProps={EditFilterDialogContentProps}
                    modalProps={updatemodelProps}
                    minWidth={500}
              >
              <div>
                <div className='ms-Grid-row'>
                  <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                    
                    <div className='Update-Title'>
                      <TextField
                        label="Title"
                        name="Title"
                        type="Text"
                        required={true}
                        onChange={(value) => 
                          this.setState({EditProjectTitle: value.target["value"]})
                        }
                        value={this.state.EditProjectTitle}
                      />
                    </div>
                      
                    <div className='Update-ClientName'>
                      <TextField
                        label="ClientName"
                        name="ClientName"
                        type="Text"
                        required={true}
                        onChange={(value) =>
                          this.setState({EditProjectClientName: value.target["value"]})
                        }
                        value={this.state.EditProjectClientName}
                      />
                    </div>

                    <div className='Update-Context'>
                     <PeoplePicker 
                          context={this.props.context}
                          titleText="People Picker"
                          personSelectionLimit={3}
                          // groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
                          showtooltip={true}
                          required={true}
                          defaultSelectedUsers={[this.state.EditProjectAssignTo.Title]}
                          // searchTextLimit={5}
                          onChange={this._getPeoplePickerItems}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={300} 
                          ensureUser={true}
                      />
                    </div>

                    <div className='Update-ProjectList'>
                      <Dropdown
                        options={(this.state.ProjectStatuslist)}
                        label="Status"
                        placeholder="Select Status"
                        required
                        defaultSelectedKey={this.state.EditProjectStatus}
                        onChange={(e, option, index) =>
                          this.setState({ EditProjectStatus: option.text})
                        }
                      />
                    </div>
                      
                    <div className='Update-Technology'> 
                      <Dropdown
                        options={(this.state.Technologylist)}
                        label="Status"
                        placeholder="Select Status"
                        required
                        defaultSelectedKey={this.state.EditProjectTechnology}
                        onChange={(e, option, index) =>
                          this.setState({ EditProjectTechnology: option.text})
                        }
                      />
                    </div>

                  </div>
                </div>
                <div className='ms-Grid-row Edit-timesheet'>
                  <div className='ms-Grid-col Update-TimeSheet'>
                      <PrimaryButton
                        text="Update"
                        type="Update"
                        iconProps={TextDocumentEdit}
                        onClick={() => this.UpdateProject(this.state.CurrenttimeSheetID , this.state.ClientID)}
                      />
                  </div>
                  <div className="ms-Grid-col Cancel-Update-TimeSheet">
                  <DefaultButton
                    type="Cancel"
                    text="Cancel"
                    iconProps={CancelIcon}  
                    onClick={() => this.setState({ EditFilterDialog: true })}
                  />
                </div>
                </div>
              </div>
              </Dialog>

              <Dialog 
                   hidden={this.state.DeleteFilterDialog}
                   onDismiss={() =>
                     this.setState({
                      DeleteFilterDialog : true,
                     })}
                   dialogContentProps={DeleteFilterDialogContentProps}
                   modalProps={deletmodelProps}
                   minWidth={300}
                >
                  <div className='Close-Icon'>
                      <Icon iconName='Cancel' onClick={() => this.setState({ DeleteFilterDialog : true })}></Icon>
                  </div>

                  <div className='Cancel-Icon'>
                          <Icon iconName='Cancel' className='cancel'/>
                  </div>

                  <div className='delete-Text'>
                    <h4>Are you sure?</h4>
                    <p>Do you really want to delete these record?</p>
                  </div>

                  <div className='ms-Grid-row'>
                    <div className='Delete-Form'>
                        <DefaultButton
                          type='Cancel'
                          text='Cancel'
                          onClick={() => this.setState({ DeleteFilterDialog : true})}
                        />

                        <PrimaryButton
                          type="Delete"
                          text="Delete"
                          onClick={() => this.DeleteProject()}
                        />
                    </div>
                  </div>  
                  
              </Dialog>

            </div>
          </div>
        </div>
    );
  }

  public async componentDidMount() {
    this.GetProject();
    this.GetClient();
    this.GetProjectItemsChoice();
  }

  public async GetProject() {
    try{
      const projectItems = await sp.web.lists.getByTitle("ProjectMaster").items.select(
        "ClientName/Title",
        "ClientName/ID",
        "AssignTo/Title",
        "AssignTo/ID",
        "ID",
        "Title",
        "Status",
        "Technology",
        "Modified",
      ).expand("ClientName","AssignTo").get().then((data) => {
        console.log(projectItems);
        console.log(data);
        this.setState({ AllProjects: data });
      });
    } catch(error) {
      console.log("Error Fetching owner data:" , error);
    }
  }

  public async GetClient() {
    const clientItems = await sp.web.lists.getByTitle("Client").items.select(
      "Title",
      "ID"
    ).get().then((data) => {
      console.log(data);
      console.log(clientItems);
      console.log(this.state.AllProjects);
    }) 
    .catch((error) => {
      console.log(error);
    });
  }

  public async GetProjectItemsChoice() {
    const choiceFiledName1 = "Status";
    const filed1 = await sp.web.lists.getByTitle("ProjectMaster").fields.getByInternalNameOrTitle(choiceFiledName1)();
    let ProjectStatuslist1 = [];
    filed1["Choices"].forEach(function (dname ,i) {
      ProjectStatuslist1.push({key: dname, text: dname });
    });
    console.log(filed1);
    this.setState({ ProjectStatuslist : ProjectStatuslist1 });

    const choiceFiledName2 = "Technology";
    const filed2 = await sp.web.lists.getByTitle("ProjectMaster").fields.getByInternalNameOrTitle(choiceFiledName2)();
    let ProjectStatuslists2 = [];
    filed2["Choices"].forEach(function (dname ,i) {
      ProjectStatuslists2.push({ key: dname, text: dname});
    });
    console.log(filed2);
    this.setState({ Technologylist : ProjectStatuslists2});
  }

  public _getPeoplePickerItems = async(items: any[]) => {

    if (items.length > 0) {
      this.setState({ AssignTo: items[0].text });
      this.setState({ AssignToID: items[0].id });
    }
    else {
      //ID=0;
      this.setState({ AssignTo: "" });
      this.setState({ AssignToID: "" });
    }
  }
    // // let assigntoid = [];
    // // const userarr = items.map(items => items.AssignToID);
    // // this.setState({ AssignToID : userarr });
    // // console.log(assigntoid);
   
    // // let userarr = [];
    // const selectedusers = items.map(items => items.id);
    // // items.forEach(user  => {
    // //   user.map({ userarr : user.Id })
    // // });
    // this.setState({ AssignToID : selectedusers[0].id });
    // console.log(this.state.AssignToID);

  public async AddProject() {
    if(this.state.Title.length == 0 || this.state.ClientName.length == 0 ||  this.state.Status.length == 0 || this.state.Technology.length == 0 ||
        this.state.AssignToID.length == 0) {
          alert("Please Complete the Details.!!");
        } 
        else {

          const clientname : any = await sp.web.lists.getByTitle("Client").items.add({
            Title: this.state.ClientName,
          }) .catch((error) => {
            console.log(error);
          });
          console.log(clientname);

            
            await sp.web.lists.getByTitle("ProjectMaster").items.add({
            Title: this.state.Title,
            ClientNameId: clientname.data.ID,
            Status: this.state.Status,
            Technology: this.state.Technology,
            AssignToId: this.state.AssignToID,
            // Modified: this.state.Modified,
          })
          .catch((error) => {
            console.log(error);
          });
          this.setState({ AddFilterDialog : true });
          this.GetProject();
         
      }
  }

  public async GetEditProject(ID) {
    let EditListForm = this.state.AllProjects.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditListForm);

    this.setState({
      EditProjectTitle: EditListForm[0].Title, 
      // ClientID: EditListForm[0].ClientID,
      EditProjectClientName : EditListForm[0].ClientName.Title,
      EditProjectStatus: EditListForm[0].Status,
      EditProjectTechnology: EditListForm[0].Technology,
      // EditProjectAssignToID: EditListForm[0].AssignToID,
      EditProjectAssignTo: EditListForm[0].AssignTo
    });
    console.log(this.state.EditProjectTitle, this.state.EditProjectClientNameId, this.state.EditProjectAssignTo, this.state.EditProjectStatus, this.state.EditProjectTechnology);
  }

  public async UpdateProject(CurrenttimeSheetID, ClientID) {
    const clientItems = await sp.web.lists.getByTitle("Client").items.getById(ClientID).update({
      Title : this.state.EditProjectClientName
    }).catch((error) => {
      console.log(error);
    });

    const updateproject = await sp.web.lists.getByTitle("ProjectMaster").items.getById(CurrenttimeSheetID).update({
      Title: this.state.EditProjectTitle,
      // ClientName: this.state.EditProjectClientNameId,
      Status: this.state.EditProjectStatus,
      Technology: this.state.EditProjectTechnology,
      AssignToId: this.state.AssignToID,       
      // AssignToID : this.state.EditProjectAssignToID
    })
    .catch ((error) => {
      console.log(error);
    });
    this.setState({ EditFilterDialog : true });
    this.GetClient(); 
    this.GetProject();
  }

  public async DeleteProject() {
    const deleteprojects = await sp.web.lists.getByTitle("ProjectMaster").items.getById(this.state.DeleteCurrentitem).delete(); 
    this.GetProject();
    this.GetClient();
    this.setState({ DeleteFilterDialog : true });
  }

}

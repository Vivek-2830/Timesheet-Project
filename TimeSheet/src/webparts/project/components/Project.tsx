import * as React from 'react';
import styles from './Project.module.scss';
import { IProjectProps } from './IProjectProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DatePicker, 
  DefaultButton, 
  Dropdown, 
  PrimaryButton, 
  TextField, 
  IIconProps, 
  Icon } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';

export interface IProjectStat {
  AllProjects: any;
  Project: any;
  Task: any;
  ClientName: any;
  Hours: any;
  TaskDescription: any;
  Date: any;
  Projectlists:any;
  ClientTitlelists: any;
  Tasklists: any;
  CurrentUserName: any;
  CurrentUserId: any;
  AddTimeSheet: any;
}

require("../assets/css/style.css");
require("../assets/css/fabric.min.css");

const CircleAdditionSolid: IIconProps = { iconName: "AddTo" };

export default class Project extends React.Component<IProjectProps,IProjectStat> {

  constructor(props: IProjectProps , state: IProjectStat){
    super(props);

    this.state = {
      AllProjects: [],
      Project: "",
      Task: "",
      ClientName: "",
      Hours: "",
      TaskDescription: "",
      Date: "",
      Projectlists:"",
      ClientTitlelists: "",
      Tasklists:"",
      CurrentUserName: "",
      CurrentUserId: "",
      AddTimeSheet : ""
      };
  }

  public render(): React.ReactElement<IProjectProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      
        <div className="project">
            <div className="ms-Grid">
              <div className="ms-Grid-row">
                <div className='ms-Grid-col ms-sm5 ms-md5 ms-lg6 ms-xl12'>
                  <div className="d-flex-header">
                    <h3 className="Title">TimeSheet</h3>
                    <div className='Add-Project'>
                      <PrimaryButton className="Add" iconProps={CircleAdditionSolid} type="Add" text='Add' onClick={() => this.handleFormAdd()} />
                    </div>
                  </div>
                </div>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-mg12 ms-lg12">
                    {
                      this.state.AllProjects.length > 0 && 
                        this.state.AllProjects.map((item, ID) => (
                          <>
                           
                            <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                                <DatePicker
                                  label='Date'
                                  placeholder="Select a date..!!"
                                  onSelectDate={(date : any) => this.handleDateChange(ID,date)
                                  }
                                  value={this.state.AllProjects[ID].Date}
                                />
                              </div>

                            <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                              <div className="Client">
                                <Dropdown
                                  placeholder="Select a Client..!!"
                                  label="Client"
                                  options={this.state.ClientTitlelists}
                                  required
                                  onChange={(e,option,index) => 
                                    this.handleClientNameChange(ID,option.key) 
                                  }
                                />
                              </div> 
                            </div>

                            <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                              <div className="Project">
                                <Dropdown
                                    label='Project'
                                    placeholder="Selct a Project..!!"
                                    options={this.state.Projectlists}
                                    required
                                    onChange={(e,option,index) =>   
                                      this.handleProjectChange(ID,option.key)
                                    }
                                />
                              </div>
                            </div>
                              
                            <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                              <div className='Task'>
                                <Dropdown
                                    label='Task'
                                    placeholder="Select a Task..!!"
                                    options={this.state.Tasklists}
                                    required
                                    onChange={(e, option, index) => 
                                      this.handleTaskChange(ID,option.text,option.key)
                                    }
                                />
                              </div>
                            </div>
                        
                            <div className='ms-Grid-col ms-sm12 ms-md1 ms-lg1 mb-10'>
                              <div className='Hour'>
                                <TextField  
                                    label='Hours'
                                    placeholder='Hours'
                                    type='Number'
                                    required={true}
                                    onChange={(value) => this.handleHoursChange(ID,value)
                                    }
                                    // value={this.state.AllProjects[ID].Hours}
                                />
                              </div>
                            </div>

                            <div className='ms-Grid-col ms-sm12 ms-md3 ms-lg3 mb-10'>
                              <div className='Task Description'>
                                <TextField
                                    label='Task Details'
                                    multiline rows = {3}
                                    required={true}
                                    onChange={(value) => this.handleTaskDescriptionChange(ID,value)
                                    }
                                    // value={this.state.AllProjects[ID].TaskDescription}
                                />
                              </div>
                            </div>
                            <div>
                                {
                                  ID == 0 ? <></> : 
                                  <>
                                  <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                    <div className='cancel-Icon'>
                                      <Icon iconName='Cancel'  onClick={() => this.handleCancelSheet(ID)}/> 
                                    </div>
                                  </div>
                                  </> 
                                }
                            </div>
                            
                          </>
                        ))
                    }
                  </div>
                      <div className='ms-Grid-row Add-Form'>
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 mt-25 text-center">
                            <div className='ms-Grid-col Submit-button'>
                              <PrimaryButton text="Submit" onClick={() => this.setState({ AddTimeSheet : ""} , () =>  this.AddTimesheet())} />
                            </div>
                            <div className='ms-Grid-col Cancel-button'>    
                                <DefaultButton text="Cancel" />
                            </div>
                          </div>
                      </div>
                </div>
              </div>
            </div>
        </div>
    );
  }

  public async componentDidMount() {
    this.GetClient();
    this.GetTimeSheetChoiceItems();
    this.GetProjectItems();
    this.handleFormAdd();
    // this.Getemployee();
  }

  public async GetClient() {
    const clientItems = await sp.web.lists.getByTitle("Client").items.select(
      "Title",
      "ID"
    ).get().then((data) => {
      let Clientdata = [];
      data.forEach(function (dname , i) {
        Clientdata.push({key: dname.ID , text: dname.Title });
      });
      console.log(clientItems);
      this.setState({ ClientTitlelists : Clientdata });
    }) 
    .catch((error) => {
      console.log(error);
    });
  }

  public async GetTimeSheetChoiceItems() {
    const choiceFiledName1 = "Task";
    const filed1 = await sp.web.lists.getByTitle("Project").fields.getByInternalNameOrTitle(choiceFiledName1)();
    let TimeSheetlists1 = [];
    filed1["Choices"].forEach(function (dname ,i) {
      TimeSheetlists1.push({key: dname, text: dname });
    });
    console.log(filed1); 
    this.setState({ Tasklists : TimeSheetlists1});
  }

  public async GetProjectItems() {
    try{
      const projectItems = await sp.web.lists.getByTitle("ProjectMaster").items.select(
        "ClientName/Title",
        "ClientName/ID",
        "Title",
        "ID"
      ).expand("ClientName").get().then((data) => {
        let projectdata = [];
          data.forEach(function (dname, i) {
            projectdata.push({ key: dname.ID ,text : dname.Title});
          });
        console.log(projectItems);
        this.setState({ Projectlists: projectdata });
      });
    } catch(error) {
      console.log("Error Fetching owner data:" , error);
    }
  }

  // public async Getemployee(){
  //   let employearr = await sp.web.currentUser.get();
  //   this.setState({ CurrentUserId : employearr.Id});
  // }

  public async handleFormAdd() {
    let employearr = await sp.web.currentUser.get();
    this.setState({ CurrentUserId : employearr.Id });
    console.log(employearr);

    let data = this.state.AllProjects;
    data.push({ Date: "",ProjectId: "",  Task: "" , TaskDescription: "", ClientNameId :"" , Hours: "", EmployeeId: this.state.CurrentUserId });
    this.setState({ AllProjects : data });
  }

  public async handleCancelSheet(ID) {
    let canceltimesheet = this.state.AllProjects;
    canceltimesheet.splice(ID , 1);
    this.setState({ AllProjects : canceltimesheet });
  }

  public async handleDateChange(ID,date) {
    let projecttimesheetdata = this.state.AllProjects;
    projecttimesheetdata[ID].Date = date;
    this.setState({ AllProjects : projecttimesheetdata });
  }

  public async handleTaskChange(ID,task,key) {
    let taskdata = this.state.AllProjects;
    taskdata[ID].Task = task;
    taskdata[ID].Task = key;
    this.setState({ AllProjects : taskdata });
  }

  public async handleTaskDescriptionChange(ID,TaskDescription) {
    let taskdescriptiondata = this.state.AllProjects;
    taskdescriptiondata[ID].TaskDescription = TaskDescription.target.value;
    this.setState({ AllProjects : taskdescriptiondata});
  }

  public async handleHoursChange(ID,Hours) {
    let hourstimesheet = this.state.AllProjects;
    hourstimesheet[ID].Hours = Hours.target.value;
    this.setState({ AllProjects : hourstimesheet });
  } 

  public async handleProjectChange(ID,key) {
    let projectsheet = this.state.AllProjects;
    // projectsheet[ID].Project = Title;
    projectsheet[ID].ProjectId = key;
    this.setState({ AllProjects : projectsheet });
  }

  public async handleClientNameChange(ID,key) {
    let clientnamesheet = this.state.AllProjects;
    // clientnamesheet[ID].ClientName = Title;
    clientnamesheet[ID].ClientNameId = key;
    this.setState({ AllProjects : clientnamesheet});
  }

  public AddTimesheet = async() => {
    let timesheet = [];
    let flag = 0;
    this.state.AllProjects.forEach(item => {
      if(item.Date.length == 0 || item.ProjectId.length == 0 || item.Task.length == 0 || item.TaskDescription.length == 0 ||
        item.ClientNameId.length == 0 || item.Hours.length == 0 ) {
          alert("Please Complete details.");
          flag = 1;
      } 
      else
      {
         if(flag == 0) {
         const timesheetadd = sp.web.lists.getByTitle("Project").items.add({
             ProjectId: item.ProjectId,
             Date : moment(item.Date).format("YYYY-MM-DDTHH:mmZ"),
             ClientNameId : item.ClientNameId,
             Task: item.Task,
             Hours: item.Hours,
             TaskDescription: item.TaskDescription,
             EmployeeId: item.EmployeeId
           }).then((response) => {
             console.log("TimeSheet Entry Added Succesfully", response);
           }).catch((error) => {
             console.log(error);
           });
           this.setState({ AddTimeSheet : timesheetadd });
           this.setState({ AddTimeSheet : ""});
          }
      }
      console.log(timesheet);
    });
  }
}



// else {
//   this.state.AllProjects.forEach((item) => {
//     sp.web.lists.getByTitle("Project").items.add(item)
//     .then(response => {
//       console.log(response);
//     }).catch(error => {
//       console.log('Error', error);
//     });
//   });
//   console.log(timesheet);
// }
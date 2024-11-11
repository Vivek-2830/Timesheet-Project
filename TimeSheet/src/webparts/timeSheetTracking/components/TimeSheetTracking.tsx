import * as React from 'react';
import styles from './TimeSheetTracking.module.scss';
import { ITimeSheetTrackingProps } from './ITimeSheetTrackingProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import $ from "jquery";
import { DetailsList, IColumn, Icon, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { saveAs } from "file-saver"; 
import * as Excel from "exceljs";

require("../assets/css/fabric.min.css");
require("../assets/css/bootstrap.min.css");
require("../assets/css/jquery.dataTables.css");
require("../assets/css/style.css");
require("../assets/Js/jquery.dataTables.js");

const Export: IIconProps = { iconName: "DownloadDocument" };

export interface ITimeSheetTrackingState{
  AllProjectData: any;
  FilterDialog: boolean;
}

let XLcolums = [
  { header: "Client", key: "ClientName" },
  { header: "Date", key: "Date" },
  { header: "Project", key: "Project" },
  { header: "Task", key: "Task" },
  { header: "Hours", key: "Hours" },
  { header: "TaskDescription", key: "TaskDescription"},
  { header: "Employee", key: "Employee"}
];

export default class TimeSheetTracking extends React.Component<ITimeSheetTrackingProps,ITimeSheetTrackingState> {

  constructor(props: ITimeSheetTrackingProps , state: ITimeSheetTrackingState){
    super(props);

    this.state = {
      AllProjectData: "",
      FilterDialog: true
    };
  }

  public render(): React.ReactElement<ITimeSheetTrackingProps> {

    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    // let tablecolumn = [
    //   {
    //     key: 'Date',
    //     name: 'Date',
    //     fieldName: 'Date',
    //     text: 'Date',
    //     isResizable: true 
    //   },
    //   {
    //     key: 'Project',
    //     name: 'Project',
    //     fieldName: 'Project',
    //     text: 'Project',
    //     isResizable: true 
    //   },
    //   {
    //     key: 'Client',
    //     name: 'Client',
    //     fieldName: 'Client',
    //     text: 'Client',
    //     isResizable: true 
    //   },
    //   {
    //     key: 'Task',
    //     name: 'Task',
    //     fieldName: 'Task',
    //     text: 'Task',
    //     isResizable: true 
    //   },
    //   {
    //     key: 'TaskDescription',
    //     name: 'TaskDescription',
    //     fieldName: 'TaskDescription',
    //     text: 'TaskDescription',
    //     isResizable: true 
    //   },
    //   {
    //     key: 'Employee',
    //     name: 'Employee',
    //     fieldName: 'Employee',
    //     text: 'Employee',
    //     isResizable: true 
    //   }
    // ]
    
   
    return (
        <div className="timeSheetTracking">
          <div className='ms-Grid'>
            
            <div className='container'>
              <h3>TimeTracking</h3>
                  <PrimaryButton
                      className='ms-Grid-col Export-button'
                      text="Export"
                      iconProps={Export}
                      onClick={() => this.ExcelData()}
                    />
                <table className='display' id="myTable">
                  <thead>
                    <tr>
                      <th>Client</th>
                      <th>Project</th>
                      <th>Date</th>
                      <th>Hours</th>
                      <th>Task</th>
                      <th>Description</th>
                      <th>Employee</th>
                    </tr>
                  </thead>
                  <tbody>
                    {
                      this.state.AllProjectData.length > 0 && (
                        this.state.AllProjectData.map((item) => {
                          return (  
                           
                            <tr>
                              <td><div><Icon iconName='AccountManagement' className='datatable-icon'></Icon> {item.ClientName}</div></td>

                              <td><div><Icon iconName='VisualsFolder' className='datatable-icon'></Icon> {item.Project}</div></td>

                              <td><div>{item.Date}</div></td>

                              <td><div>{item.Hours}</div></td>

                              <td><div>{item.Task}</div></td>

                              <td><p className='description'>{item.TaskDescription}</p></td>
                               <td><div><Icon iconName='People' className='datatable-icon'></Icon> {item.Employee}</div></td>
                            </tr>
                          );
                        })
                      )
                    }
                  </tbody>
                </table>
            </div>
          </div>
        </div>
    );
  }

  public async componentDidMount() {
    this.GetProject();
  }

  // public GetProject =  async() => {
  //   try {
  //     const projectitems = await sp.web.lists.getByTitle("Project").items.select(
  //       'ClientName/Title', 
  //       'ClientName/ID', 
  //       'Title', 
  //       'ID', 
  //       'Task', 
  //       'Hours', 
  //       'TaskDescription',
  //       'Date', 
  //       'Project/ID', 
  //       'Project/Title', 
  //       'Employee/ID',
  //       'Employee/Title').expand('ClientName', 'Project', 'Employee').get().then((data)=> {
  //       console.log(projectitems);
  //       console.log(data);
  //       this.setState({ AllProjectData: projectitems })
          
  //     });
  //     $('#myTable').DataTable(
  //     );
  //     console.log(this.state.AllProjectData);
  //   }
  //    catch(error) {
  //     console.log(error);
  //   }
  // }

  public GetProject = async () => {
    await sp.web.lists.getByTitle("Project").items.select('ClientName/Title', 'ClientName/ID', 'Title', 'ID', 'Task', 'Hours', 'TaskDescription',
      'Date', 'Project/ID', 'Project/Title', 'Employee/ID', 'Employee/Title').expand('ClientName', 'Project', 'Employee').get().then((data) => {
        let AllData = [];
        console.log(data);
 
        if (data.length > 0) {
          data.forEach((item) => {
            AllData.push({
              Date: item.Date ? item.Date.split("T")[0] : "",
              Project: item.Project.Title ? item.Project.Title : "",
              ClientName: item.ClientName.Title ? item.ClientName.Title : "",
              Task: item.Task ? item.Task : "",
              Hours: item.Hours ? item.Hours : "",
              TaskDescription: item.TaskDescription ? item.TaskDescription : "",
              Employee: item.Employee.Title ? item.Employee.Title : "",
            });
          });
          this.setState({ AllProjectData: AllData },
            () => {
              $('#myTable').DataTable(
                );
            }
            );
          console.log(this.state.AllProjectData);
        }
 
      }).catch((err) => {
        console.log(err);
      });
  }
 
  public async ExcelData() {
    // const fileName = 'TimeSheet.xls'
    // const exportoExcel =() =>{}
    
      const web = sp.web;
      const siteTitle = await web.select("Title").get();
  
      const workbook = new Excel.Workbook();
  
      if (this.state.AllProjectData.length > 0) {
        try {
          const fileName =
            moment().format("DD/MM/YYYY HH:MM") +
            " TimeSheet Excel Overview " +
            siteTitle.Title;
          const worksheet = workbook.addWorksheet();
  
          
          worksheet.columns = XLcolums;
  
        
          worksheet.getRow(1).font = { bold: true, color: { argb: "00000000" } };
          worksheet.getRow(1).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ffffffff" }, 
          };
  
         
          worksheet.columns = [
            { width: 70 },
            { width: 30 },
            { width: 15 },
            { width: 50 },
          ];
          worksheet.getColumn(1).alignment = {
            horizontal: "left",
            wrapText: true,
            vertical: "middle",
          };
          worksheet.getColumn(2).alignment = {
            horizontal: "left",
            wrapText: true,
            vertical: "middle",
          };
          worksheet.getColumn(3).alignment = {
            horizontal: "left",
            wrapText: true,
            vertical: "middle",
          };
          worksheet.getColumn(4).alignment = {
            horizontal: "left",
            wrapText: true,
            vertical: "middle",
          };
  
          const oddRowColor = "FFFFFF"; 
          const evenRowColor = "fbfbfb";
          const borderColor = "aaaaaa"; 
  
         
          this.state.AllProjectData.forEach((singleData: any, index: number) => {
            const row = worksheet.addRow(singleData);
  
          
            const fillColor =
              index % 2 === 0 ? { argb: oddRowColor } : { argb: evenRowColor };
            row.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: fillColor,
            };
  
            
            row.eachCell((cell, colNumber) => {
              cell.border = {
                top: { style: "thin", color: { argb: borderColor } },
                left: { style: "thin", color: { argb: borderColor } },
                bottom: { style: "thin", color: { argb: borderColor } },
                right: { style: "thin", color: { argb: borderColor } },
              };
            });
          });
  
          
          const buf = await workbook.xlsx.writeBuffer();
  
          saveAs(new Blob([buf]), `${fileName}.xlsx`);
        } catch (error) {
          console.error("Something Went Wrong", error.message);
        }
      } else {
        alert(
          "Please select News you want to export, then click the 'Export' button."
        );
      }
    }
}

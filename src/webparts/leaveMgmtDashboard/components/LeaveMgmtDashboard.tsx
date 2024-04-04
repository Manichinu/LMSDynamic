import * as React from 'react';
// import styles from './LeaveMgmtDashboard.module.scss';
import { ILeaveMgmtDashboardProps } from './ILeaveMgmtDashboardProps';
import { ILeaveMgmtDashboardState } from './ILeaveMgmtDashboardState';
import { escape } from '@microsoft/sp-lodash-subset';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'jquery/dist/jquery.min.js';
import { _SiteGroups } from '@pnp/sp/site-groups/types';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
//Datatable Modules

import "datatables.net-dt/js/dataTables.dataTables";
import "datatables.net-dt/css/jquery.dataTables.min.css";
import "datatables.net";
//import "datatables.net-buttons";
//import "datatables.net-dt/css/jquery.dataTables.css";
//import "datatables.net-buttons-dt/css/buttons.dataTables.css";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/presets/all";
import "@pnp/sp/fields";
import swal from "sweetalert";
import { Web } from '@pnp/sp/webs';
import "@pnp/sp/site-users/web";
import * as $ from 'jquery';
import { sp } from "@pnp/sp";
import { SPComponentLoader } from "@microsoft/sp-loader";
import * as moment from 'moment';
import Swal from "sweetalert2";
import { DateTimeFieldFormatType, FieldTypes } from "@pnp/sp/fields/types";
import LeaveMgmt from './LeaveMgmt';
import Holiday from './Holiday';
import PermissionDashboard from './PermissionDashboard';
import PermissionRequest from './PermissionRequest';
import Aboutus from './Aboutus';
import "../css/style"

var AttachmentCopies = [];
var spfxdatatable = null;
let CausalArr = [];
let CausalArrtotal = [];
let SickArr = [];
var firstpage = "First";
var Lastpage = "Last";
var Nextpage = "Next";
var Previouspage = "Previous";
let ItemId;
var CurrentUSERNAME = "";
var Usertype = "";
var datesInRange: any = [];
var InBetweenDates: any = [];
var updateDateIdNo: any;
var TotalDaysLeaveApplied: any;
var LeaveTypee: any;
var LeaveStatuss: any;
var SpecificDate: any;

// const NewWeb = Web('${this.props.siteurl}/');
var NewWeb: any;
let progressEndValue = 100;
let overAllValue = 0;

export interface LeaveMgmtDashboardState {
  DatatableItems: any[];
  IsAdmin: boolean;
  CurrentUserName: string;
  CurrentUserDesignation: string;
  CurrentUserProfilePic: string;
  Gender: string;
  IsUser: boolean;
  Empemail: string;
  CurrentUserId: number;
  LeaveBalanceItems: any[];
  LeaveMgmtDashboard: boolean;
  Holiday: boolean;
  LeaveMgmt: boolean;
  PermissionDashboard: boolean;
  PermissionRequest: boolean;
  AboutUs: boolean;
  Configure: boolean;
}

export default class LeaveMgmtDashboard extends React.Component<ILeaveMgmtDashboardProps, LeaveMgmtDashboardState> {
  public constructor(props: ILeaveMgmtDashboardProps) {
    super(props);
    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
    );
    SPComponentLoader.loadCss(`https://fonts.googleapis.com`);
    SPComponentLoader.loadCss(`https://fonts.gstatic.com" crossorigin`);
    SPComponentLoader.loadCss(
      `https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet`
    );
    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css`
    );

    SPComponentLoader.loadScript(
      `https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js`
    );

    SPComponentLoader.loadScript(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.js`
    );
    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
    );

    // SPComponentLoader.loadCss(
    //   `${this.props.siteurl}/SiteAssets/LeavePortal/css/style.css?v=1.14`
    // );


    sp.setup({
      spfxContext: this.props.context
    });

    this.state = {
      DatatableItems: [],
      IsAdmin: false,

      CurrentUserName: "",
      CurrentUserDesignation: "",
      CurrentUserProfilePic: "",
      Gender: "",
      IsUser: false,
      CurrentUserId: null,
      LeaveBalanceItems: [],
      Empemail: "",
      LeaveMgmtDashboard: false,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: false,
      Configure: false,
    };
    NewWeb = Web("" + this.props.siteurl + "")

  }

  public logout() {
    // window.location.href = `https://login.microsoftonline.com/556e6b1f-b49d-4278-8baf-db06eeefc8e9/oauth2/v2.0/logout`;
    //localStorage.clear();
    window.location.href = `https://login.windows.net/common/oauth2/logout`;
  }

  public GetCurrentUserDetails() {
    var reacthandler = this;
    $.ajax({
      url: `${reacthandler.props.siteurl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        var email = resultData.d.Email;
        var Name = resultData.d.DisplayName;
        var Designation = resultData.d.Title;
        var gender = resultData.d.Streetaddress;
        reacthandler.setState({
          CurrentUserName: Name,
          CurrentUserDesignation: Designation,
          Empemail: email,
          CurrentUserProfilePic: `${reacthandler.props.siteurl}/_layouts/15/userphoto.aspx?size=l&username=${email}`,
          Gender: gender
          // },() =>{this.GetleaveBalance()
        });
        reacthandler.GetleaveBalance(email); //Get User data from list and bind it in form 
        // reacthandler.Checkuserexists();
        // reacthandler.checkUserInGroup("LMS Admin");
      },




      error: function (jqXHR, textStatus, errorThrown) {

      }

    });

  }

  public async componentDidMount() {

    this.GetCurrentUserDetails();
    this.checkConfiguredOrNot();
    // NewWeb.lists.getByTitle("BalanceCollection").fields.addCalculated("My Field", {
    //   Formula: "=1+1",
    //   Group: "MyGroup"
    // });
    // this.checkIfListExists();
    // this.createSitePage();
    // this.createGroup();
    // this.createLeaveRequestList();
    // this.createLeaveCancellationHistoryList();
    // this.createEmployeePermissionList();
    // this.createBalanceCollectionList();
    // this.createHolidayCollectionList();
    // this.createLeaveTypeCollectionList();
    // NewWeb.lists.getByTitle("BalanceCollection").fields.getByTitle("CasualLeaveBalance").update({ Formula: "=1+1" },);

    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    let userDetails = await this.spLoggedInUser(this.props.context);

    let userID = userDetails.Id;
    this.setState({ CurrentUserId: userID });

    await this.isOwnerGroupMember();
    //this.loadTable();
    //await this.GetListitems();

  }
  public async checkConfiguredOrNot() {
    NewWeb.lists.getByTitle("Configure Master").items.get().then((items: any) => {
      if (items.length != 0) {
        this.setState({
          LeaveMgmtDashboard: true
        })
        $("#header-section").show()
      }
    }).catch((error: any) => {
      this.setState({
        Configure: true,
      })
      console.error("An error occurred:", error);
    });
  }
  public async createAllDynamicLists() {
    try {
      $("#configure").hide();
      $(".progress_container").show();
      await this.configureListCreation();
      await this.createSitePage();
      await this.createGroup();
      await this.createLeaveRequestList();
      await this.createLeaveCancellationHistoryList();
      await this.createEmployeePermissionList();
      await this.createBalanceCollectionList();
      await this.createHolidayCollectionList();
      await this.createLeaveTypeCollectionList();

    } catch (error) {
      console.error("Error configuring lists:", error);
    }
  }
  public async checkIfListExists() {
    try {
      // Retrieve all lists from the site
      const lists = await NewWeb.lists.get();
      var listNames = ["LeaveRequest", "Leave Cancellation History", "EmployeePermission", "BalanceCollection", "HolidayCollection", "LeaveTypeCollection"]
      listNames.map((item: any) => {
        const listExists = lists.some((list: any) => list.Title === item);
        if (listExists) {
          console.log(`The list "${item}" exists in the site.`);
        } else {
          console.log(`The list "${item}" does not exist in the site.`);
        }
      })

    } catch (error) {
      console.error("Error:", error);
    }
  }
  public async createSitePage() {
    try {
      // Create a new client side page
      var pageTitle = "LeaveManagement"
      const page = await sp.web.addClientsidePage(pageTitle, "Page Title", "Article");;
      // Set the content of the page
      // await page.save(pageContent);
      console.log(`Site Page "${pageTitle}" created successfully.`);
    } catch (error) {
      console.error("Error:", error);
    }
  }
  public async createGroup() {
    var groupName = "LMS Admin";
    var groupDescription = "Description"
    try {
      // Create a new group
      await sp.web.siteGroups.add({
        Title: groupName,
        Description: groupDescription
      });

      console.log(`Group "${groupName}" created successfully.`);
    } catch (error) {
      console.error("Error:", error);
    }
  }
  public async createLeaveRequestList() {
    var ListName = "LeaveRequest";
    var batch = NewWeb.createBatch();
    var Columns = [
      { Name: "Day", Type: "SingleLine" },
      { Name: "Time", Type: "SingleLine" },
      { Name: "Reason", Type: "MultiLine" },
      { Name: "Status", Type: "SingleLine" },
      { Name: "Requester", Type: "SingleLine" },
      { Name: "Approver", Type: "SingleLine" },
      { Name: "EmployeeEmail", Type: "SingleLine" },
      { Name: "Days", Type: "Number" },
      { Name: "ManagerComments", Type: "SingleLine" },
      { Name: "AppliedDate", Type: "SingleLine" },
      { Name: "StartDate", Type: "SingleLine" },
      { Name: "EndDate", Type: "SingleLine" },
      { Name: "LeaveType", Type: "SingleLine" },
      { Name: "RequestSessionMasterID", Type: "SingleLine" },
      { Name: "ApproverEmail", Type: "SingleLine" },
      { Name: "CompOff", Type: "MultiLine" },

    ]
    await NewWeb.lists.add(ListName).then(() => {
      Columns.map(async (item: any) => {
        if (item.Type == "SingleLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addText(item.Name, 255, {
            Group: "Custom Column",
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "MultiLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addMultilineText(item.Name, 255, true, false, false, false, {
            Group: "Custom Column",
            RichText: false,
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Number") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
      })
      // Execute the batch
      batch.execute().then(function () {
        console.log("Batch operations completed successfully for creating " + ListName + " list");
      }).catch(function (error: any) {
        console.log("Error in batch operations for creating " + ListName + " list: " + error);
      });
    })
  }
  public async createLeaveCancellationHistoryList() {
    var ListName = "Leave Cancellation History";
    var batch = NewWeb.createBatch();
    var Columns = [
      { Name: "Day", Type: "SingleLine" },
      { Name: "Time", Type: "SingleLine" },
      { Name: "Reason", Type: "MultiLine" },
      { Name: "Status", Type: "SingleLine" },
      { Name: "Requester", Type: "SingleLine" },
      { Name: "Approver", Type: "SingleLine" },
      { Name: "EmployeeEmail", Type: "SingleLine" },
      { Name: "Days", Type: "Number" },
      { Name: "ManagerComments", Type: "SingleLine" },
      { Name: "AppliedDate", Type: "SingleLine" },
      { Name: "StartDate", Type: "SingleLine" },
      { Name: "EndDate", Type: "SingleLine" },
      { Name: "LeaveType", Type: "SingleLine" },
      { Name: "RequestSessionMasterID", Type: "SingleLine" },
      { Name: "ApproverEmail", Type: "SingleLine" },
      { Name: "CompOff", Type: "MultiLine" },

    ]
    await NewWeb.lists.add(ListName).then(() => {
      Columns.map(async (item: any) => {
        if (item.Type == "SingleLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addText(item.Name, 255, {
            Group: "Custom Column",
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "MultiLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addMultilineText(item.Name, 255, true, false, false, false, {
            Group: "Custom Column",
            RichText: false,
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Number") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
      })
      // Execute the batch
      batch.execute().then(function () {
        console.log("Batch operations completed successfully for creating " + ListName + " list");
      }).catch(function (error: any) {
        console.log("Error in batch operations for creating " + ListName + " list: " + error);
      });
    })
  }
  public async createEmployeePermissionList() {
    var ListName = "EmployeePermission";
    var batch = NewWeb.createBatch();
    var Columns = [
      { Name: "Approver", Type: "SingleLine" },
      { Name: "Status", Type: "SingleLine" },
      { Name: "PermissionHour", Type: "SingleLine" },
      { Name: "Reason", Type: "MultiLine" },
      { Name: "TimeUpto", Type: "SingleLine" },
      { Name: "EmployeeEmail", Type: "SingleLine" },
      { Name: "PermissionOn", Type: "SingleLine" },
      { Name: "timefromwhen", Type: "SingleLine" },
      { Name: "ManagerComments", Type: "MultiLine" },
      { Name: "Requester", Type: "SingleLine" },
      { Name: "ApproverEmail", Type: "SingleLine" },
      { Name: "CancelledBy", Type: "SingleLine" },

    ]
    await NewWeb.lists.add(ListName).then(() => {
      Columns.map(async (item: any) => {
        if (item.Type == "SingleLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addText(item.Name, 255, {
            Group: "Custom Column",
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "MultiLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addMultilineText(item.Name, 255, true, false, false, false, {
            Group: "Custom Column",
            RichText: false,
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Number") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
      })
      // Execute the batch
      batch.execute().then(function () {
        console.log("Batch operations completed successfully for creating " + ListName + " list");
      }).catch(function (error: any) {
        console.log("Error in batch operations for creating " + ListName + " list: " + error);
      });
    })
  }
  public async createBalanceCollectionList() {
    var ListName = "BalanceCollection";
    var batch = NewWeb.createBatch();
    var Columns = [
      { Name: "EmployeeEmail", Type: "SingleLine" },
      { Name: "Year", Type: "SingleLine" },
      { Name: "SickLeave", Type: "Number" },
      { Name: "SickLeaveUsed", Type: "Number" },
      { Name: "OtherLeave", Type: "Number" },
      { Name: "EmployeeName", Type: "SingleLine" },
      { Name: "CasualLeave", Type: "Number" },
      { Name: "CasualLeaveUsed", Type: "Number" },
      { Name: "OtherLeaveUsed", Type: "Number" },
      { Name: "MaternityLeave", Type: "Number" },
      { Name: "MaternityLeaveUsed", Type: "Number" },
      { Name: "PaternityLeave", Type: "Number" },
      { Name: "PaternityLeaveUsed", Type: "Number" },
      { Name: "SickLeaveBalance", Type: "Calculated", Formula: "=(SickLeave-SickLeaveUsed)" },
      { Name: "OtherLeaveBalance", Type: "Calculated", Formula: "=(OtherLeave-OtherLeaveUsed)" },
      { Name: "PaternityLeaveBalance", Type: "Calculated", Formula: "=(PaternityLeave-PaternityLeaveUsed)" },
      { Name: "MaternityLeaveBalance", Type: "Calculated", Formula: "=(MaternityLeave-MaternityLeaveUsed)" },
      { Name: "CasualLeaveBalance", Type: "Calculated", Formula: "=(CasualLeave-CasualLeaveUsed)" },
      { Name: "EarnedLeave", Type: "Number" },
      { Name: "EarnedLeaveUsed", Type: "Number" },
      { Name: "EarnedLeaveBalance", Type: "Calculated", Formula: "=(EarnedLeave-EarnedLeaveUsed)" },
      { Name: "EmployeeID", Type: "SingleLine" },
      { Name: "StartDate", Type: "SingleLine" },
      { Name: "EndDate", Type: "SingleLine" },
      { Name: "Manager", Type: "Person" },



    ]
    await NewWeb.lists.add(ListName).then(() => {
      Columns.map(async (item: any) => {
        if (item.Type == "SingleLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addText(item.Name, 255, {
            Group: "Custom Column",
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "MultiLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addMultilineText(item.Name, 255, true, false, false, false, {
            Group: "Custom Column",
            RichText: false,
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Number") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Calculated") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addCalculated(item.Name)
            .then(async () => {
              const progress = (1 * 100 / 74);
              this.updateProgress(progress);
              // NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
              // NewWeb.lists.getByTitle(ListName).fields.getByTitle(item.Name).update({ Formula: item.Formula },);
              await Promise.all([
                NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name),
                NewWeb.lists.getByTitle(ListName).fields.getByTitle(item.Name).update({ Formula: item.Formula })
              ]);
            }).then(() => {
              console.log(`${item.Name} column created successfully`)
            })
        }
        else if (item.Type == "Person") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addUser(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
      })
      // Execute the batch
      batch.execute().then(function () {
        console.log("Batch operations completed successfully for creating " + ListName + " list");
      }).catch(function (error: any) {
        console.log("Error in batch operations for creating " + ListName + " list: " + error);
      });
    })
  }
  public async createHolidayCollectionList() {
    var ListName = "HolidayCollection";
    var batch = NewWeb.createBatch();
    var Columns = [
      { Name: "StartDate", Type: "Date" },
      { Name: "HolidayName", Type: "SingleLine" },
      { Name: "Holidaytype", Type: "SingleLine" },

    ]
    await NewWeb.lists.add(ListName).then(() => {
      Columns.map(async (item: any) => {
        if (item.Type == "SingleLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addText(item.Name, 255, {
            Group: "Custom Column",
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Date") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addDateTime(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }

      })
      // Execute the batch
      batch.execute().then(function () {
        console.log("Batch operations completed successfully for creating " + ListName + " list");
      }).catch(function (error: any) {
        console.log("Error in batch operations for creating " + ListName + " list: " + error);
      });
    })
  }
  public async createLeaveTypeCollectionList() {
    var ListName = "LeaveTypeCollection";
    var batch = NewWeb.createBatch();
    var Columns = [
      { Name: "Types", Type: "SingleLine" },
      { Name: "Details", Type: "MultiLine" },

    ]
    await NewWeb.lists.add(ListName).then(() => {
      Columns.map(async (item: any) => {
        if (item.Type == "SingleLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addText(item.Name, 255, {
            Group: "Custom Column",
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "MultiLine") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addMultilineText(item.Name, 255, true, false, false, false, {
            Group: "Custom Column",
            RichText: false,
          }).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 74);
            this.updateProgress(progress);
          })
        }
      })
      // Execute the batch
      batch.execute().then(function () {
        console.log("Batch operations completed successfully for creating " + ListName + " list");
      }).catch(function (error: any) {
        console.log("Error in batch operations for creating " + ListName + " list: " + error);
      });
    })
  }
  public async configureListCreation() {
    try {
      const listTitle = "Configure Master";
      const listDescription = "Form Template";
      NewWeb.lists.add(listTitle, listDescription, 100, false).then(() => {
        console.log(`${listTitle} List created successfully`);
        NewWeb.lists.getByTitle(listTitle).items.add({
          Title: "Configured"
        })
      });
    } catch (error) {
      console.error("Error creating list:", error);
    }
  }
  public updateProgress(value: any) {
    overAllValue += value;
    console.log("Progress", overAllValue)
    if (overAllValue >= progressEndValue) {
      $(".progress-value").text(`100%`);
      Swal.fire('Configured successfully!', '', 'success').then(() => {
        location.reload();
      })
      $(".progress_container").hide();
    } else {
      $(".progress-value").text(`${Math.ceil(overAllValue)}%`);
      $(".circular-progress").css("background", `conic-gradient(#7d2ae8 ${Math.ceil(overAllValue) * 3.6}deg, #ededed 0deg)`);
    }
  }
  public Checkuserexists() {
    if (this.checkUserInGroup("LMS Admin")) {

      // setTimeout(() => {

      this.setState({ IsAdmin: true });
      //this.GetAdminlistitems();

      //    }, 1000);


    } else {

      // setTimeout(() => {
      // console.log("Not Exists");
      this.setState({ IsAdmin: false });
      // this.GetUserlistitems();

      // }, 1000);
    }
  }

  public checkUserInGroup(strGroup: string) {

    let InGroup: boolean = false;
    const title = (_SiteGroups as any)['Title'];

    // let grp:any = sp.web.currentUser.groups.get().then((r: any) => {
    //   r.forEach((grp: _SiteGroups) => {
    //     if (grp["Title"] == strGroup) {
    //       InGroup = true;

    //       this.GetAdminlistitems();


    //     }
    //     else {
    //       InGroup = false;
    //       this.GetUserlistitems();
    //     }

    //     console.log(grp["Title"]);
    //   });
    // });

    return InGroup;

  }
  /** Check if the current user is in Owners group **/
  public async isOwnerGroupMember() {
    var reacthandler = this;
    let userDetails = await this.spLoggedInUser(this.props.context);

    let userID = userDetails.Id;
    console.log(userID);
    $.ajax({

      // url: `${reacthandler.props.siteurl}/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq  + ${this.props.userId}`,
      url: `${reacthandler.props.siteurl}/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq ${userID}`,

      type: "GET",

      headers: { 'Accept': 'application/json; odata=verbose;' },

      success: function (resultData) {

        if (resultData.d.results.length == 0) {
          console.log("User not in group : LMS Admin Owners");
          setTimeout(() => {
            reacthandler.GetUserlistitems();
          }, 1000);


        } else {

          console.log("User in group : LMS Admin Owners");

          setTimeout(() => {

            reacthandler.GetAdminlistitems();
          }, 1000);

        }

      },

      error: function (jqXHR, textStatus, errorThrown) {
        console.log("Error while checking user in Owner's group");
      }

    });

  }
  public async Checkuserforlogin() {
    let userDetails = await this.spLoggedInUser(this.props.context);
    console.log(userDetails.Id);
    let currusername = userDetails.LoginName;
    let groups = await NewWeb.currentUser.groups();

    for (var i = 0; i < groups.length; i++) {
      if (groups[i].Title == "LMS Admin") {


        this.setState({ IsAdmin: true });
        await this.GetAdminlistitems();

      } else {
        // this.setState({ IsAdmin: false });
        await this.GetUserlistitems();
      }

    }
  }
  public GetUserlistitems() {
    var reactHandler = this;
    NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "*", "StartDate", "EndDate", "Reason", "Days", "Requester", "EmployeeEmail", "Day", "LeaveType", "Status", "AppliedDate", "CompOff").filter(`Author/Id eq ${this.props.userId}`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

      // await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Day", "Reason", "Days", "Requester", "EmployeeEmail", "LeaveType", "Status", "AppliedDate").filter(`Author/Id eq ${this.props.userId}`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

      .then((items: any) => {
        if (items.length != 0) {

          reactHandler.setState({
            DatatableItems: items
          });
          this.loadTable();

        }
        else {
          this.loadTable();
        }
      });
  }
  public GetAdminlistitems() {
    this.setState({ IsAdmin: true });
    var reactHandler = this;
    NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "*", "StartDate", "EndDate", "Reason", "Days", "Requester", "EmployeeEmail", "Day", "LeaveType", "Status", "AppliedDate", "CompOff").expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

      // await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Day", "Reason", "Days", "Requester", "EmployeeEmail", "LeaveType", "Status", "AppliedDate").filter(`Author/Id eq ${this.props.userId}`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

      .then((items: any) => {
        if (items.length != 0) {

          reactHandler.setState({
            DatatableItems: items
          });
          this.loadTable();

        }
        else {
          this.loadTable();
        }
      });
  }
  /*Get Current Logged In User*/

  public async GetListitems() {
    var reacthandler = this;
    var UserType = "";
    let groups = await NewWeb.currentUser.groups();
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].Title == 'LMS Admin') {

        this.setState({ IsAdmin: true });
        UserType = "Admin";
        Usertype = UserType;
        //reacthandler.Checkusertype(UserType);


        return false;
      }
      else {
        UserType = "User";
        Usertype = UserType;

        // reacthandler.Checkusertype(UserType);

      }
    }
    console.log(UserType);
    {/* if (UserType == "User") {
      console.log("User :" + UserType);
      let userDetails = await this.spLoggedInUser(this.props.context);
      console.log(userDetails.Id);
      let userID = userDetails.Id;

      await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Day", "Reason", "Days", "Requester", "EmployeeEmail", "LeaveType", "Status", "AppliedDate").filter(`Author/Id eq  ${userID}`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

        //  await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "startdate", "enddate","Day", "Reason", "Days", "Requester", "EmployeeEmail", "leavetype","Status","AppliedDate").filter(`Author/Id eq ${this.state.CurrentUserId}`).orderBy("Created", false).top(5000).get()
        .then((items) => {
          if (items.length != 0) {

            reactHandler.setState({
              DatatableItems: items
            });
            this.loadTable();

          }
          else {
            this.loadTable();
          }
        });

    } else {
      if (UserType =="Admin") {
      console.log("Admin :" + UserType);
      await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Reason", "Days", "Requester", "EmployeeEmail", "Day", "LeaveType", "Status", "AppliedDate").expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

        .then((items) => {
          if (items.length != 0) {

            reactHandler.setState({

              DatatableItems: items
            });
            this.loadTable();

          }
          else {
            this.loadTable();
          }
        });
      }
    }*/}

  }

  public async GetListitemsnew() {
    var reactHandler = this;
    var UserType = "";
    let groups = await NewWeb.currentUser.groups();
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].Title == 'LMS Admin') {
        UserType = "Admin";
        // this.setState({ IsAdmin: true });
        console.log(UserType);
        Usertype = UserType;

        return Usertype;

      } else {

        UserType = "User";
        Usertype = "User"

      }
    }

    return Usertype;
    /* if (UserType == "User") {
   
     //  await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Day", "Reason", "Days", "Requester", "EmployeeEmail", "LeaveType", "Status", "AppliedDate").filter(`Author/Id eq ${this.state.CurrentUserId}`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()
       await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Day", "Reason", "Days", "Requester", "EmployeeEmail", "LeaveType", "Status", "AppliedDate").filter("EmployeeEmail eq '" +this.state.Empemail +"'").expand('AttachmentFiles').orderBy("Created", false).top(5000).get()
       .then((items) => {
           if (items.length != 0) {
   
             reactHandler.setState({
               DatatableItems: items
             });
             this.loadTable();
   
           }
           else {
             this.loadTable();
           }
         });
   
     } else {
       //if (UserType =="Admin") {
   
       await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Reason", "Days", "Requester", "EmployeeEmail", "Day", "LeaveType", "Status", "AppliedDate").expand('AttachmentFiles').orderBy("Created", false).top(5000).get()
   
         .then((items) => {
           if (items.length != 0) {
   
             reactHandler.setState({
   
               DatatableItems: items
             });
             this.loadTable();
   
           }
           else {
             this.loadTable();
           }
         });*/


  }

  public async Checkusertype(UserType: string) {
    var reactHandler = this;
    if (UserType == "User") {
      console.log("User :" + UserType);
      let userDetails = await this.spLoggedInUser(this.props.context);
      console.log(userDetails.Id);
      let userID = userDetails.Id;

      await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "*", "StartDate", "EndDate", "Day", "Reason", "Days", "Requester", "EmployeeEmail", "LeaveType", "Status", "AppliedDate", "CompOff").filter(`Author/Id eq ${this.props.userId}`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()
        //  await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate", "Day", "Reason", "Days", "Requester", "EmployeeEmail", "LeaveType", "Status", "AppliedDate").filter("EmployeeEmail eq '" + this.state.Empemail + "'").expand('AttachmentFiles').orderBy("Created", false).top(5000).get()
        .then((items: any) => {
          if (items.length != 0) {

            reactHandler.setState({
              DatatableItems: items
            });
            this.loadTable();

          }
          else {
            this.loadTable();
          }
        });

    } else {
      //if (UserType =="Admin") {
      console.log("Admin :" + UserType);
      await NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "*", "StartDate", "EndDate", "Reason", "Days", "Requester", "EmployeeEmail", "Day", "LeaveType", "Status", "AppliedDate", "CompOff").expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

        .then((items: any) => {
          if (items.length != 0) {

            reactHandler.setState({

              DatatableItems: items
            });
            this.loadTable();

          }
          else {
            this.loadTable();
          }
        });

    }
  }

  private async spLoggedInUser(ctx: any) {
    try {
      const web = Web(ctx.pageContext.site.absoluteUrl);
      return await web.currentUser.get();
    } catch (error) {
      console.log("Error in spLoggedInUserDetails : " + error);
    }
  }


  public async GetGroupmembers() {
    var userGroups = []

    // Gets the associated visitors group of a web
    const visitorGroup = await sp.web.associatedVisitorGroup();
    console.log(visitorGroup);
    // Gets the associated members group of a web
    const memberGroup = await sp.web.associatedMemberGroup();
    // get all groups the current user belongs to
    return sp.web.currentUser.groups().then(function (groups) {
      for (var i = 0; i < groups.length; i++) {
        userGroups.push(groups[i].Title);
      }
    });
  }
  /*public async GetGroupMembers()
  {
    let groups = await NewWeb.currentUser.groups();
    graph.groups.top(999).select('mailNickname,id').get().then(groups=>{
      console.log(groups);
      groups.forEach( group=>{
   
        if(group['Internal Tmax Guestuser']== "Internal Tmax Guestuser")
        {
           let groupId = group['id'];
           const groupMembers =  graph.groups.getById(groupId).expand("members")().then(group=>{
            console.log(group.members);
          });
   
        }
      })
    }).catch((err) => {
      console.log("Error fetching Group ID "+err)});
   
  }*/





  public GetCurrentUserName() {
    var reactHandler = this;
    $.ajax({
      url: `${this.props.siteurl}/_api/web/currentUser`,
      method: "GET",
      headers: {
        Accept: "application/json; odata=verbose",
      },
      success: function (data) {
        CurrentUSERNAME = data.d.Title;
        reactHandler.setState({
          CurrentUserName: CurrentUSERNAME

        });
      },
      error: function (data) { },
    });
  }

  public loadTable() {
    //  ($('#LMSDashboard') as any).DataTable.destroy();
    // ($('#LMSDashboard') as any).DataTable({
    //   pageLength: 5,
    //   "bSort": false,
    //   "bDestroy": true,

    //   lengthMenu: [[5, 10, 20, 50, -1], [5, 10, 20, 50, "All"]],

    //   initComplete: function () {

    //     this.api().columns().every(function () {

    //       var column = this;

    //       var select = $('<select><option value="">All</option></select>')

    //         .appendTo($(column.header()).empty()).on('change', function () {

    //           var val = ($ as any).fn.dataTable.util.escapeRegex(

    //             ($(this) as any).val()

    //           );

    //           column.search(val ? '^' + val + '$' : '', true, false).draw();


    //         });

    //       column.data().unique().sort().each(function (d, j) {
    //         // select.append('<option value="' + d + '</option>')
    //         var temp2 = d;
    //         if (temp2.indexOf(">") != -1) {
    //           var temp = d.split(">");
    //           var temporary = temp[3].split("<");

    //           select.append('<option value="' + temporary[0] + '">' + temporary[0] + '</option>')
    //         } else {
    //           select.append('<option value="' + d + '">' + d + '</option>')

    //         }



    //       });


    //     });

    //   }

    // });


    {/*            column.data().unique().sort().each(function (d, j) {

              select.append('<option value="' + d + '">' + d + '</option>')
  
            });
  
          });
  
        }
  
      });*/}


    $.fn.dataTable.ext.errMode = "none";

    $("#LMSDashboard").DataTable({
      ordering: false,
      pageLength: 5,
      lengthMenu: [
        [5, 10, 20, 50, 100, -1],
        [5, 10, 20, 50, 100, "All"],
      ],
      dom: "Blfrtip",

      initComplete: function () {
        this.api()
          .columns()
          .every(function () {
            var column = this;
            var select = $('<select><option value="">All</option></select>')
              .appendTo($(column.header()).empty())
              .on("change", function () {
                var val = $.fn.dataTable.util.escapeRegex(
                  ($(this) as any).val()
                );

                column.search(val ? "^" + val + "$" : "", true, false).draw();
              });

            column
              .data()
              .unique()
              .sort()
              .each(function (d: string, j: any) {
                select.append('<option value="' + d + '">' + d + "</option>");
              });
          });
      },
    });

  }

  public GetleaveBalance(email: any) {
    CausalArr = [];
    CausalArrtotal = [];
    SickArr = [];
    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1;
    console.log(currentYear);

    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "CasualLeave", "CasualLeaveBalance", "EmployeeEmail", "CasualLeaveUsed", "SickLeave", "SickLeaveUsed", "SickLeaveBalance", "OtherLeaveBalance", "OtherLeave", "OtherLeaveUsed", "EarnedLeaveBalance", "EarnedLeave", "EarnedLeaveUsed", "PaternityLeaveUsed", "PaternityLeave", "PaternityLeaveBalance", "MaternityLeave", "MaternityLeaveBalance", "MaternityLeaveUsed", "StartDate", "EndDate").filter(`EmployeeEmail eq '${email}'`).get()

      .then((items: any) => {

        if (items.length != 0) {
          console.log(items)
          this.setState({

            LeaveBalanceItems: items
          });


        }
      });



    {/*var url = `https://tmxin.sharepoint.com/sites/POC/SPIP/_api/web/lists/getbytitle('BalanceCollection')/items?$select=CasualLeave,CasualLeaveBalance,EmployeeEmail,CasualLeaveUsed,SickLeave,SickLeaveUsed,SickLeaveBalance,OtherLeaveBalance,OtherLeave,OtherLeaveUsed,EarnedLeaveBalance,EarnedLeave,EarnedLeaveUsed,PaternityLeaveUsed,PaternityLeave,PaternityLeaveBalance,MaternityLeave,MaternityLeaveBalance,MaternityLeaveUsed&$filter('Author/EmployeeEmail eq '${this.state.email}'')`;
  $.ajax({
    url: url,
    type: "GET",
    async: false,
    headers: { 'Accept': 'application/json; odata=verbose;' },
    success: function (resultData) {
      console.log(resultData);

      reactHandler.setState({

        LeaveBalanceItems: resultData.d.results
      });


      /*  for (var i = 0; i < resultData.d.results.length; i++) {
          if (resultData.d.results[i].CasualLeaveUsed == "CasualLeaveUsed") {
            CausalArr.push(resultData.d.results[i]);
          }
                
        }
   
        
   //  var TotalCL = `${CausalArr.length}/${this.state.TotalCasualleave}`
         var TotalCL = `${CausalArr.length}/${12}`
          $("#casualLeave").html(TotalCL);
    },
    error: function (jqXHR, textStatus, errorThrown) {
    }
  });*/}

  }
  public GetLeaveDetails() {
    var reactHandler = this;
    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('LeaveRequest')/items?$select=StartDate,EndDate,LeaveType,Day,Reason,Requester,AppliedDate,Status&$orderby=Created desc`;
    $.ajax({
      url: url,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        console.log(resultData);
        reactHandler.setState({

          DatatableItems: resultData.d.results
        });

        setTimeout(() => {
          reactHandler.loadTable();
        }, 1000);

        {/* for (var i = 0; i < resultData.d.results.length; i++) {
          if (resultData.d.results[i].LeaveType == "AnualLeave") {
            AnnualArr.push(resultData.d.results[i]);


          }
        }

        var TotalAnualLeave = `${AnnualArr.length}/${ttlAnnulLeave}`
      $("#EarnedLeave").html(TotalAnualLeave)*/}
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }

  public Displayleaveform() {

    location.href = `${this.props.siteurl}/SitePages/LeaveApplication.aspx?env=WebView`;
  }

  public Cancel_Request_(itemidno: number, totalDays: number, StartDate: moment.MomentInput, EndDate: moment.MomentInput, LeaveType: any, LeaveStatus: any, items: any) {
    var startdate = moment(StartDate).format('DD-MM-YYYY')
    var endtdate = moment(EndDate).format('DD-MM-YYYY')
    var currentdate = moment().format('DD-MM-YYYY')
    console.log(items)
    // if (endtdate == currentdate || endtdate > currentdate) {
    if (totalDays == 0.5 || totalDays == .5 || totalDays == 1) {
      swal({
        title: ` "Are you sure?"`,
        text: "Would you like to cancel the leave?",
        icon: "warning",
        buttons: ["No", "Yes"],
        dangerMode: true,
      } as any).then((willdelete) => {
        if (willdelete) {
          NewWeb.lists.getByTitle("LeaveRequest").items.getById(itemidno).update({
            Status: "Cancelled"
          }).then(() => {
            NewWeb.lists.getByTitle("Leave Cancellation History").items.add({
              LeaveType: items.LeaveType,
              Day: items.Day,
              Time: items.Time,
              StartDate: items.StartDate,
              EndDate: items.EndDate,
              Reason: items.Reason,
              Requester: items.Requester,
              AppliedDate: items.AppliedDate,
              Days: items.Days,
              EmployeeEmail: items.EmployeeEmail,
              RequestSessionMasterID: items.RequestSessionMasterID,
              Approver: items.Approver,
              ApproverEmail: items.ApproverEmail,
              CompOff: items.CompOff,
              ManagerComments: items.ManagerComments,
              Status: "Cancelled",
              CancelledBy: this.state.CurrentUserName

            })
            this.Get_Blance_Count(totalDays, LeaveType, LeaveStatus)

          })



        }
      })




    }

    // }

  }






  public Get_Blance_Count(totalDays: any, leavetype: any, LeaveStatus: any) {


    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "EmployeeEmail").filter(`EmployeeEmail eq '${this.state.Empemail}'`).get()
      .then((results: any) => {
        if (results.length != 0) {


          this.Update_Blance_Count(results, totalDays, leavetype, LeaveStatus)

        }
      }).then(() => {
        swal({
          text: "Leave cancelled successfully!",
          icon: "success",
        }).then(() => {
          // location.href = "${this.props.siteurl}/SitePages/Dashboard.aspx?env=WebView";
          location.reload();
        });
      })

  }


  // public Update_Blance_Count(result: any[], totaldaysapplied_leave: number, leavetype: string) {
  //   //Earned Leave 
  //   //Casual Leave 
  //   //Sick Leave 
  //   //Unpaid Leave 
  //   //Maternity Leave 
  //   //Paternity Leave
  //   if (totaldaysapplied_leave == 0.5 || totaldaysapplied_leave == 1 || totaldaysapplied_leave == .5)
  //     if (leavetype == "Casual Leave") {
  //       var totalCasual_leave: number = result[0].CasualLeave
  //       var casualleaveused: number = result[0].CasualLeaveUsed + totaldaysapplied_leave
  //       var CasualLeave_blance: number = totalCasual_leave - casualleaveused

  //       NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
  //         CasualLeaveUsed: casualleaveused,
  //         CasualLeaveBalance: CasualLeave_blance
  //       })

  //     } else if (leavetype == "Earned Leave") {
  //       //EarnedLeaveUsed EarnedLeave EarnedLeaveBalance
  //       var total_Earned_Leave: number = result[0].EarnedLeave
  //       var Earned_leave_used = result[0].EarnedLeaveUsed + totaldaysapplied_leave
  //       var Earned_Leave_blance: number = total_Earned_Leave - Earned_leave_used;

  //       NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
  //         EarnedLeaveUsed: Earned_leave_used,
  //         EarnedLeaveBalance: Earned_Leave_blance
  //       })

  //     } else if (leavetype == "Sick Leave") {
  //       //SickLeave SickLeaveUsed SickLeaveBalance

  //       var total_sick_Leave: number = result[0].SickLeave
  //       var sick_leave_used = result[0].SickLeaveUsed + totaldaysapplied_leave
  //       var sick_blance: number = total_sick_Leave - sick_leave_used;

  //       NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
  //         SickLeaveUsed: sick_leave_used,
  //         SickLeaveBalance: sick_blance
  //       })
  //     } else if (leavetype == "Unpaid Leave") {
  //       //OtherLeaveUsed OtherLeaveBalance OtherLeave

  //       var total_unpaid_Leave: number = result[0].OtherLeave
  //       var unpaid_Leave: number = result[0].OtherLeaveUsed + totaldaysapplied_leave
  //       var unpaid_Leave_blance: number = total_unpaid_Leave - unpaid_Leave;

  //       NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
  //         OtherLeaveUsed: unpaid_Leave,
  //         OtherLeaveBalance: unpaid_Leave_blance
  //       })
  //     } else if (leavetype == "Maternity Leave") {
  //       //MaternityLeaveBalance MaternityLeaveUsed MaternityLeave
  //       var total_MaternityLeave: number = result[0].MaternityLeave
  //       var MaternityLeave_used: number = result[0].MaternityLeaveUsed + totaldaysapplied_leave
  //       var MaternityLeaveBalance: number = total_MaternityLeave - MaternityLeave_used;
  //       NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
  //         MaternityLeaveBalance: MaternityLeaveBalance,
  //         MaternityLeaveUsed: MaternityLeave_used
  //       })
  //     } else if (leavetype == "Paternity Leave") {
  //       //PaternityLeaveBalance PaternityLeaveUsed PaternityLeave


  //       var Total_PaternityLeave: number = result[0].PaternityLeave;
  //       var PaternityLeave_used: number = result[0].PaternityLeaveUsed + totaldaysapplied_leave;

  //       var PaternityLeave_Balance: number = Total_PaternityLeave - PaternityLeave_used;
  //       NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
  //         PaternityLeaveBalance: PaternityLeave_Balance,
  //         PaternityLeaveUsed: PaternityLeave_used
  //       })
  //     }

  // }
  public Update_Blance_Count(result: any[], totaldaysapplied_leave: number, leavetype: string, LeaveStatus: any) {
    //Earned Leave 
    //Casual Leave 
    //Sick Leave 
    //Unpaid Leave 
    //Maternity Leave 
    //Paternity Leave
    if (LeaveStatus != "Pending") {
      if (totaldaysapplied_leave == 0.5 || totaldaysapplied_leave == 1 || totaldaysapplied_leave == .5 || totaldaysapplied_leave >= 1) {
        if (leavetype == "Casual Leave") {
          var casualleaveused: number = result[0].CasualLeaveUsed - totaldaysapplied_leave

          NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
            CasualLeaveUsed: casualleaveused,
          })

        } else if (leavetype == "Earned Leave") {
          //EarnedLeaveUsed EarnedLeave EarnedLeaveBalance
          var Earned_leave_used = result[0].EarnedLeaveUsed - totaldaysapplied_leave

          NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
            EarnedLeaveUsed: Earned_leave_used,
          })

        } else if (leavetype == "Sick Leave") {
          //SickLeave SickLeaveUsed SickLeaveBalance

          var sick_leave_used = result[0].SickLeaveUsed - totaldaysapplied_leave

          NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
            SickLeaveUsed: sick_leave_used,
          })
        } else if (leavetype == "Unpaid Leave") {
          //OtherLeaveUsed OtherLeaveBalance OtherLeave

          var unpaid_Leave: number = result[0].OtherLeaveUsed - totaldaysapplied_leave

          NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
            OtherLeaveUsed: unpaid_Leave,
          })
        } else if (leavetype == "Maternity Leave") {
          //MaternityLeaveBalance MaternityLeaveUsed MaternityLeave
          var MaternityLeave_used: number = result[0].MaternityLeaveUsed - totaldaysapplied_leave
          NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
            MaternityLeaveUsed: MaternityLeave_used
          })
        } else if (leavetype == "Paternity Leave") {
          //PaternityLeaveBalance PaternityLeaveUsed PaternityLeave


          var PaternityLeave_used: number = result[0].PaternityLeaveUsed - totaldaysapplied_leave;

          NewWeb.lists.getByTitle("BalanceCollection").items.getById(result[0].ID).update({
            PaternityLeaveUsed: PaternityLeave_used
          })
        }
      }
    }

  }
  public Cancel_Request_or_change_LeaveDate(itemidno: any, totalDays: any, StartDate: any, EndDate: any, LeaveType: any, LeaveStatus: any, items: any) {
    // var startdate = moment(StartDate).format('DD-MM-YYYY')
    // var endtdate = moment(EndDate).format('DD-MM-YYYY')
    // var currentdate = moment().format('DD-MM-YYYY')


    swal({
      title: `Are you sure?`,
      text: "Would you like to cancel the leave?",
      icon: "warning",
      buttons: ["No", "Yes"],
      dangerMode: true,
    } as any).then((willdelete) => {
      if (willdelete) {
        if (LeaveStatus != "Pending") {
          Swal.fire({
            icon: "warning",
            showDenyButton: true,
            showCancelButton: true,
            confirmButtonText: "Cancel specific leave date",
            denyButtonText: `Cancel this leave`,
            cancelButtonText: 'Close',
            customClass: {
              container: 'cancel-popup',
            },
          }).then((result) => {
            /* Read more about isConfirmed, isDenied below */
            if (result.isConfirmed) {
              Swal.fire({
                title: "<p>Select Date</p>",
                html: "<input type='date' id='cancelation_date' />",
                confirmButtonText: "Submit",
                customClass: {
                  container: 'cancel-date',
                },
                showCloseButton: true,
                allowOutsideClick: true,
                preConfirm: () => {
                  var selectedDate = $("#cancelation_date").val();
                  if (selectedDate == "") {
                    Swal.showValidationMessage("Please select a date");
                  }
                  return selectedDate;
                },
              }).then((result) => {
                if (result.isConfirmed) {
                  var CurrentDate = moment().format("YYYY-MM-DD")
                  var SelectedDate = $("#cancelation_date").val()
                  if (SelectedDate != "") {
                    if (CurrentDate != SelectedDate) {
                      this.updateLeaveDates(SelectedDate)
                    } else {
                      swal({
                        text: "Don't select the current date",
                        icon: "error"
                      });
                    }
                  }
                }
              });



              //this.Change_leave_Date_or_Cancel_Leave(totalDays, LeaveType)
              // Parse the start and end dates
              updateDateIdNo = itemidno;
              TotalDaysLeaveApplied = totalDays;
              LeaveTypee = LeaveType;
              LeaveStatuss = LeaveStatus;
              SpecificDate = items
              console.log(SpecificDate)
              var startDate = new Date(StartDate);
              var endDate = new Date(EndDate);
              $('#cancelation_date').attr('min', StartDate);
              $('#cancelation_date').attr('max', EndDate);

              // Array to store the dates in between
              datesInRange = [];
              InBetweenDates = [];
              // Iterate through the dates and add them to the array
              for (var currentDate = startDate; currentDate <= endDate; currentDate.setDate(currentDate.getDate() + 1)) {
                // Format the date as "YYYY-MM-DD" and push to the array
                var formattedDate = currentDate.toISOString().split('T')[0];
                datesInRange.push(formattedDate);
              }

              // console.log(datesInRange);
              InBetweenDates = datesInRange.slice(1, -1);
            } else if (result.isDenied) {
              NewWeb.lists.getByTitle("LeaveRequest").items.getById(itemidno).update({
                Status: "Cancelled"
              }).then(() => {
                NewWeb.lists.getByTitle("Leave Cancellation History").items.add({
                  LeaveType: items.LeaveType,
                  Day: items.Day,
                  Time: items.Time,
                  StartDate: items.StartDate,
                  EndDate: items.EndDate,
                  Reason: items.Reason,
                  Requester: items.Requester,
                  AppliedDate: items.AppliedDate,
                  Days: items.Days,
                  EmployeeEmail: items.EmployeeEmail,
                  RequestSessionMasterID: items.RequestSessionMasterID,
                  Approver: items.Approver,
                  ApproverEmail: items.ApproverEmail,
                  CompOff: items.CompOff,
                  ManagerComments: items.ManagerComments,
                  Status: "Cancelled",
                  CancelledBy: this.state.CurrentUserName

                })
              })
              this.Get_Blance_Count(totalDays, LeaveType, LeaveStatus)
            }
          });
        } else {
          NewWeb.lists.getByTitle("LeaveRequest").items.getById(itemidno).update({
            Status: "Cancelled"
          }).then(() => {
            NewWeb.lists.getByTitle("Leave Cancellation History").items.add({
              LeaveType: items.LeaveType,
              Day: items.Day,
              Time: items.Time,
              StartDate: items.StartDate,
              EndDate: items.EndDate,
              Reason: items.Reason,
              Requester: items.Requester,
              AppliedDate: items.AppliedDate,
              Days: items.Days,
              EmployeeEmail: items.EmployeeEmail,
              RequestSessionMasterID: items.RequestSessionMasterID,
              Approver: items.Approver,
              ApproverEmail: items.ApproverEmail,
              CompOff: items.CompOff,
              ManagerComments: items.ManagerComments,
              Status: "Cancelled",
              CancelledBy: this.state.CurrentUserName

            })
          })
          this.Get_Blance_Count(totalDays, LeaveType, LeaveStatus)
        }

        // swal({
        //   title: ``,
        //   text: "",
        //   icon: "warning",
        //   buttons: ["Cancel specific leave date", "Cancel this leave"],
        //   dangerMode: true,
        //   closeOnClickOutside: false,
        //   showCloseButton: true,
        //   showCancelButton: true
        // } as any).then((willdelete) => {
        //   if (willdelete) {

        //     NewWeb.lists.getByTitle("LeaveRequest").items.getById(itemidno).update({
        //       Status: "Cancelled"
        //     }).then(() => {
        //       NewWeb.lists.getByTitle("Leave Cancellation History").items.add({
        //         LeaveType: items.LeaveType,
        //         Day: items.Day,
        //         Time: items.Time,
        //         StartDate: items.StartDate,
        //         EndDate: items.EndDate,
        //         Reason: items.Reason,
        //         Requester: items.Requester,
        //         AppliedDate: items.AppliedDate,
        //         Days: items.Days,
        //         EmployeeEmail: items.EmployeeEmail,
        //         RequestSessionMasterID: items.RequestSessionMasterID,
        //         Approver: items.Approver,
        //         ApproverEmail: items.ApproverEmail,
        //         CompOff: items.CompOff,
        //         ManagerComments: items.ManagerComments,
        //         Status: "Cancelled"
        //       })
        //     })
        //     this.Get_Blance_Count(totalDays, LeaveType, LeaveStatus)

        //     // 
        //   } else {
        //     // $(".popup_show").show()
        //     Swal.fire({
        //       title: "<p>Select Date</p>",
        //       html: "<input type='date' id='cancelation_date' />",
        //       confirmButtonText: "Submit",
        //       customClass: {
        //         container: 'cancel-date',
        //       },
        //       showCloseButton: true,
        //       allowOutsideClick: true,
        //       preConfirm: () => {
        //         var selectedDate = $("#cancelation_date").val();
        //         if (selectedDate == "") {
        //           Swal.showValidationMessage("Please select a date");
        //         }
        //         return selectedDate;
        //       },
        //     }).then((result) => {
        //       if (result.isConfirmed) {
        //         var CurrentDate = moment().format("YYYY-MM-DD")
        //         var SelectedDate = $("#cancelation_date").val()
        //         if (SelectedDate != "") {
        //           if (CurrentDate != SelectedDate) {
        //             this.updateLeaveDates(SelectedDate)
        //           } else {
        //             swal({
        //               text: "Don't select the current date",
        //               icon: "error"
        //             });
        //           }
        //         }
        //       }
        //     });



        //     //this.Change_leave_Date_or_Cancel_Leave(totalDays, LeaveType)
        //     // Parse the start and end dates
        //     updateDateIdNo = itemidno;
        //     TotalDaysLeaveApplied = totalDays;
        //     LeaveTypee = LeaveType;
        //     LeaveStatuss = LeaveStatus;
        //     SpecificDate = items
        //     console.log(SpecificDate)
        //     var startDate = new Date(StartDate);
        //     var endDate = new Date(EndDate);
        //     $('#cancelation_date').attr('min', StartDate);
        //     $('#cancelation_date').attr('max', EndDate);

        //     // Array to store the dates in between
        //     datesInRange = [];
        //     InBetweenDates = [];
        //     // Iterate through the dates and add them to the array
        //     for (var currentDate = startDate; currentDate <= endDate; currentDate.setDate(currentDate.getDate() + 1)) {
        //       // Format the date as "YYYY-MM-DD" and push to the array
        //       var formattedDate = currentDate.toISOString().split('T')[0];
        //       datesInRange.push(formattedDate);
        //     }

        //     // console.log(datesInRange);
        //     InBetweenDates = datesInRange.slice(1, -1);
        //     // console.log(InBetweenDates)
        //   }


        // })
      }
    })


  }
  public dateValidation(SelectedDate: any) {
    var FormStatus = true;
    var Date = SelectedDate
    InBetweenDates.map((item: any) => {
      if (item == Date) {
        FormStatus = false;
      }
    })
    return FormStatus;
  }
  public updateLeaveDates(SelectedDate: any) {
    if (this.dateValidation(SelectedDate)) {
      var Date = SelectedDate
      var LeaveDates = datesInRange.filter((date: any) => date != Date)
      var StartDate = LeaveDates[0];
      var EndDate = LeaveDates[LeaveDates.length - 1];

      NewWeb.lists.getByTitle("LeaveRequest").items.getById(updateDateIdNo).update({
        StartDate: StartDate,
        EndDate: EndDate,
        Days: LeaveDates.length
      }).then(() => {
        NewWeb.lists.getByTitle("Leave Cancellation History").items.add({
          LeaveType: SpecificDate.LeaveType,
          Day: SpecificDate.Day,
          Time: SpecificDate.Time,
          StartDate: SelectedDate,
          EndDate: SelectedDate,
          Reason: SpecificDate.Reason,
          Requester: SpecificDate.Requester,
          AppliedDate: SpecificDate.AppliedDate,
          Days: 1,
          EmployeeEmail: SpecificDate.EmployeeEmail,
          RequestSessionMasterID: SpecificDate.RequestSessionMasterID,
          Approver: SpecificDate.Approver,
          ApproverEmail: SpecificDate.ApproverEmail,
          CompOff: SpecificDate.CompOff,
          ManagerComments: SpecificDate.ManagerComments,
          Status: "Cancelled",
          CancelledBy: this.state.CurrentUserName

        })
        var ReduceLeaveDays = TotalDaysLeaveApplied - LeaveDates.length
        this.Get_Blance_Count(ReduceLeaveDays, LeaveTypee, LeaveStatuss)

      })

    } else {
      // console.log("Don't select inbetween dates")
      swal({
        text: "Don't select inbetween date",
        icon: "error"
      });
    }
  }
  // public Change_leave_Date_or_Cancel_Leave(totaldays_leave: any, LeaveType: any) {
  //   let currentYear = new Date().getFullYear()
  //   NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "EmployeeEmail").filter(`EmployeeEmail eq '${this.state.Empemail}' and Year eq ${currentYear}`).get()
  //     .then((results) => {
  //       if (results.length != 0) {

  //       }
  //     })

  // }
  public showLeaveMgmtDashboard() {
    this.setState({
      LeaveMgmtDashboard: true,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: false,
    })
  }
  public showHoliday() {
    this.setState({
      LeaveMgmtDashboard: false,
      Holiday: true,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: false,
    })
  }
  public showLeaveMgmt() {
    this.setState({
      LeaveMgmtDashboard: false,
      Holiday: false,
      LeaveMgmt: true,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: false,
    })
  }
  public showPermissionDashboard() {
    this.setState({
      LeaveMgmtDashboard: false,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: true,
      PermissionRequest: false,
      AboutUs: false,
    })
  }
  public showPermissionRequest() {
    this.setState({
      LeaveMgmtDashboard: false,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: true,
      AboutUs: false,
    })
  }
  public showAboutus() {
    this.setState({
      LeaveMgmtDashboard: false,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: true,
    })
  }
  public render(): React.ReactElement<ILeaveMgmtDashboardProps> {
    let count = 0;
    let handler = this;

    const CasualLeaveBodycontent: JSX.Element[] = this.state.LeaveBalanceItems.map(function (item, key) {

      count++;
      return (

        <span>{(item.CasualLeaveUsed)}/{parseFloat(item.CasualLeaveBalance)}</span>

      );
    });
    const SickLeaveBodycontent: JSX.Element[] = this.state.LeaveBalanceItems.map(function (item, key) {

      count++;
      return (

        <span>{(item.SickLeaveUsed)}/{parseFloat(item.SickLeaveBalance)}</span>

      );
    });
    const EarnedLeaveBodycontent: JSX.Element[] = this.state.LeaveBalanceItems.map(function (item, key) {

      count++;
      return (

        <span>{(item.EarnedLeaveUsed)}/{parseFloat(item.EarnedLeaveBalance)}</span>

      );
    });
    const OtherLeaveBodycontent: JSX.Element[] = this.state.LeaveBalanceItems.map(function (item, key) {

      count++;
      return (

        <span>{(item.OtherLeaveUsed)}</span>

      );
    });
    const MaternityLeaveBodycontent: JSX.Element[] = this.state.LeaveBalanceItems.map(function (item, key) {

      count++;
      return (

        <span>{(item.MaternityLeaveUsed)}/{parseFloat(item.MaternityLeaveBalance)}</span>

      );
    });
    const PaternityLeaveBodycontent: JSX.Element[] = this.state.LeaveBalanceItems.map(function (item, key) {

      count++;
      return (

        <span>{(item.PaternityLeaveUsed)}/{parseFloat(item.PaternityLeaveBalance)}</span>

      );
    });
    const DataTableBodycontent: JSX.Element[] = this.state.DatatableItems.map(function (item, key) {

      count++;
      return (

        <tr id={`${key}-row-id`}>
          <td>{key + 1}</td>
          {handler.state.IsAdmin == true &&
            <td>{item.Requester}</td>
          }
          <td>{moment(item.AppliedDate).format('DD-MM-YYYY')}</td>
          <td>{item.LeaveType}</td>
          <td>{moment(item.StartDate).format('DD-MM-YYYY')}</td>
          <td>{moment(item.EndDate).format('DD-MM-YYYY')}</td>
          <td className="reason-td">{item.Reason}</td>

          <td>{item.Day}</td>
          <td>{item.CompOff === "" || item.CompOff === null || item.CompOff === undefined ? "-" : item.CompOff}</td>

          {item.Status == "Pending" ?
            <td className="status pending text-center">{item.Status}</td>
            :
            item.Status == "Approved" ?
              <td className="status approved text-center">{item.Status}</td>
              :
              item.Status == "Rejected" ?
                <td className="status rejected text-center">{item.Status}</td>
                :
                (item.Status == "Cancelled" || item.Status == "Cancel") ?
                  <td className="status rejected text-center">{item.Status}</td>
                  :
                  <></>

          }

          <td><ul>
            {item.AttachmentFiles && item.AttachmentFiles.map(function (afile: any, key: any) {
              return (

                <li><a href={`${afile.ServerRelativeUrl}`} data-interception="off" target="_blank">{afile.FileName}</a></li>
              )
            })}
          </ul></td>
          {/*  {handler.state.IsAdmin == true &&
            <td>{item.Requester}</td>
          }*/}

          <td style={{ cursor: "pointer" }} className='cancel-section'>

            {(item.State !== "Cancel" || item.State !== "Cancelled") &&
              <>
                {((item.Days <= 1) && handler.state.Empemail == item.EmployeeEmail && item.Status != "Cancelled" && item.Status != "Rejected" && moment(item.EndDate, "YYYY-MM-DD").isAfter(moment(), 'day')) &&


                  <p onClick={() => handler.Cancel_Request_(item.Id, item.Days, item.StartDate, item.EndDate, item.LeaveType, item.Status, item)}><img src={require("../img/cancel.svg")} alt="image" /></p>
                }

                {((item.Days > 1) && handler.state.Empemail == item.EmployeeEmail && item.Status != "Cancelled" && item.Status != "Rejected" && moment(item.EndDate, "YYYY-MM-DD").isAfter(moment(), 'day')) &&


                  <p onClick={() => handler.Cancel_Request_or_change_LeaveDate(item.Id, item.Days, item.StartDate, item.EndDate, item.LeaveType, item.Status, item)}> <img src={require("../img/cancel.svg")} alt="image" /></p>
                }

              </>
            }
          </td>
        </tr>

      );
    });
    return (
      <>
        {this.state.Configure == true &&
          <div className='config'>
            <div className="progress_container" style={{ display: "none" }}>
              <div className="circular-progress">
                <span className="progress-value">0%</span>
              </div>
            </div>
            <button type="button" id='configure' className="btn btn-primary" onClick={() => this.createAllDynamicLists()} >Click here to Configure</button>
          </div>
        }
        <div id='header-section' style={{ display: "none" }}>
          <header>
            <div className="clearfix container-new">
              <div className="logo">
                <img src={require("../img/logo_small.png")} alt="image" />
              </div>
              <div className="header-title"><h3>Leave Management System</h3></div>
              <div className="notification-part">
                <ul>
                  <li className="person-details">
                    {/*<span id="CurrentUser-Profilepicture"> <img src={`${this.state.CurrentUserProfilePic}`} alt="image" /> <span>  </span>  </span>*/}
                    <span id="CurrentUser-displayname">{this.state.CurrentUserName}</span>
                    <a href="#" onClick={this.logout}><img src={require("../img/logout.png")} /></a>
                  </li>
                </ul>
              </div>
            </div>
          </header>
          <nav>
            <ul>
              <li className={this.state.LeaveMgmtDashboard == true ? "active" : ""} onClick={() => this.showLeaveMgmtDashboard()}>Home  </li>
              <li className={this.state.AboutUs == true ? "active" : ""} onClick={() => this.showAboutus()}> About  </li>
              <li className={this.state.Holiday == true ? "active" : ""} onClick={() => this.showHoliday()}> Holidays  </li>
              <li className={this.state.PermissionDashboard == true ? "active" : ""} onClick={() => this.showPermissionDashboard()} id='permission-dashboard'> Permission  </li>
            </ul>
          </nav>
        </div>
        {this.state.LeaveMgmtDashboard == true &&
          <div>
            <div className="container">
              <div className="dashboard-wrap">
                {/* <div>
              <button onClick={this.logout}>Logout</button>
           </div>*/}
                <div className="tab-headings clearfix">
                  <ul className="nav nav-pills">
                    <li className="active"><a data-toggle="pill" href="#home">Dashboard</a></li>

                    {/* <li><a data-toggle="pill" href="#menu1">Calender</a></li>
                <li><a data-toggle="pill" href="#menu2">Department</a></li>*/}
                  </ul>
                  {/* <td><a href="#" onClick={() => handler.View(item.Id)}>View</a></td>*/}
                  {this.state.IsAdmin == true &&

                    <a href="${this.props.siteurl}/Lists/LeaveRequest/Approvedlist.aspx" className="btn btn-outline leave-req-link " id="submit">View leave list</a>
                  }
                  <button className="btn btn-outline" id="submit" onClick={() => this.showLeaveMgmt()}> New Leave Request  </button>

                </div>

                <div className="tab-content">
                  <div id="home" className="tab-pane fade in active">
                    <div className="three-blocks-wrap">

                      <div className="row">
                        <div className="col-md-4"><a href="">
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/approved.png")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Earned Leave </p>

                              {EarnedLeaveBodycontent}
                            </div>

                          </div> </a>
                        </div>
                        <div className="col-md-4"><a href="">
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/pending.png")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Casual Leave </p>

                              {CasualLeaveBodycontent}

                            </div>

                          </div></a>
                        </div>
                        <div className="col-md-4"><a href="">
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/sick leave.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Sick Leave </p>
                              {SickLeaveBodycontent}
                            </div>
                          </div></a>
                        </div>
                        <div className="col-md-4"><a href="">
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/maternity.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Maternity Leave </p>
                              {MaternityLeaveBodycontent}
                            </div>
                          </div></a>
                        </div>
                        <div className="col-md-4"><a href="">
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/paternity leave.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Paternity Leave </p>
                              {PaternityLeaveBodycontent}
                            </div>
                          </div></a>
                        </div>
                        <div className="col-md-4"><a href="">
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/other leave.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Unpaid Leave </p>
                              {OtherLeaveBodycontent}
                            </div>
                          </div></a>
                        </div>
                      </div>

                    </div>
                    <div className="table-wrap">
                      <div className="table-search-wrap clearfix">
                        <div className="table-search relative">
                          {/* <input type="text" placeholder="Search Here" className="" />
                      <img src="https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/img/search.svg" alt="image" />*/}
                        </div>

                      </div>
                      <table id="LMSDashboard" className="table" >
                        <thead>
                          <tr>
                            <th></th>

                            {this.state.IsAdmin == true &&
                              <th></th>
                            }
                            <th></th>
                            <th></th>
                            <th></th>
                            <th></th>
                            <th className="reason-select-input"></th>
                            <th></th>
                            <th></th>
                            <th className="text-center"></th>
                            <th></th>
                            <th></th>

                          </tr>
                        </thead>
                        <thead>

                          <tr>
                            <th>S.No</th>
                            {this.state.IsAdmin == true &&
                              <th>Employee Name</th>
                            }
                            <th>Requested On</th>
                            <th>Leave Type</th>
                            <th>Start Date</th>
                            <th>End Date</th>
                            <th className="reason-td">Reason</th>
                            <th>Day</th>
                            <th>Compensation Date</th>
                            <th className="text-center"> Status  </th>
                            <th>Attachments</th>
                            <th>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {DataTableBodycontent}
                        </tbody>

                      </table>
                    </div>

                  </div>

                </div>

                {/* <div style={{ display: "none" }} className="popup_show leave_overlay">
              <div id="input-cancel-form" className="overlay_popup_comment">
                <div className="clearfix">
                  <img title="close" src="https://etccgov.sharepoint.com/sites/CompanyApplication/SiteAssets/ETCC/IMAGES/img/close.svg" style={{ float: "right" }} className="cancel_email_btn" /></div>
                <div className="create_details">
                  <div className="row">
                    <div className="col-md-3">
                      <label>  Select Date  <i className="required">*</i> </label>
                      <input type='date' id='cancelation_date' className="form-control start-date" autoComplete='off'></input>

                    </div>
                    
                  </div>
                </div>
                <div className="create_btn">
                  <button className="submit_btn reject-btn" onClick={() => this.updateLeaveDates()}>Submit</button>
                </div>
              </div>
            </div> */}
              </div>
            </div>
          </div>
        }
        {this.state.Holiday == true &&
          <Holiday description={''} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />
        }
        {this.state.LeaveMgmt == true &&
          <LeaveMgmt description={''} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.PermissionDashboard == true &&
          <PermissionDashboard description={''} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.PermissionRequest == true &&
          <PermissionRequest description={''} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.AboutUs == true &&
          <Aboutus description={''} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
      </>
    );
  }

}

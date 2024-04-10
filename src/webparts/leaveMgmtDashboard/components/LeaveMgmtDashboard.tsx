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
import ApprovalDashboard from './Approvals';
import PermissionApprovalDashboard from './PermissionApprovals';
import "../css/style";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";

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

var NewWeb: any;
let progressEndValue = 100;
let overAllValue = 0;
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 1000;

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
  Approvals: boolean;
  PermissionApprovalDashboard: boolean;
  IsCurrentUserisManager: boolean;
  leaveType: string;
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
      Approvals: false,
      PermissionApprovalDashboard: false,
      IsCurrentUserisManager: false,
      leaveType: ""
    };
    NewWeb = Web("" + this.props.siteurl + "")

  }

  public logout() {
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
        reacthandler.GetleaveBalance(email);
      },




      error: function (jqXHR, textStatus, errorThrown) {

      }

    });

  }
  public async componentDidMount() {
    this.GetCurrentUserDetails();
    this.checkConfiguredOrNot();

    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");


    let userDetails = await this.spLoggedInUser(this.props.context);

    let userID = userDetails.Id;
    this.setState({ CurrentUserId: userID });

    await this.isOwnerGroupMember();
    this.GetApproverConfigurationItems()

  }
  public async GetApproverConfigurationItems() {
    await NewWeb.lists.getByTitle("Approver Configuration").items.select("*").getAll()
      .then((items: any) => {
        items.map((item: any) => {
          if (this.state.CurrentUserId == item.ApproverId) {
            this.setState({
              IsCurrentUserisManager: true
            })
            return
          }
        })
        console.log("ManagerItems", items)
      });
    const searchParams = new URLSearchParams(window.location.search);
    const hasTab = searchParams.has("tab");
    if (hasTab) {
      var tabName = searchParams.get("tab");
      if (tabName == "permission") {
        this.showPermissionApprovalsDashboard()
      } else if (tabName == "leave") {
        this.showApprovalsDashboard()
      }

    }
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
      await this.addCurrentUserDetails();
      await this.createApproverConfigurationList();
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
            const progress = (1 * 100 / 75);
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
            const progress = (1 * 100 / 75);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Number") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 75);
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
            const progress = (1 * 100 / 75);
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
            const progress = (1 * 100 / 75);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Number") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 75);
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
            const progress = (1 * 100 / 75);
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
            const progress = (1 * 100 / 75);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Number") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 75);
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
  // public async createBalanceCollectionList() {
  //   var ListName = "BalanceCollection";
  //   var batch = NewWeb.createBatch();
  //   var Columns = [
  //     { Name: "EmployeeEmail", Type: "SingleLine" },
  //     { Name: "Year", Type: "SingleLine" },
  //     { Name: "SickLeave", Type: "Number" },
  //     { Name: "SickLeaveUsed", Type: "Number" },
  //     { Name: "OtherLeave", Type: "Number" },
  //     { Name: "EmployeeName", Type: "SingleLine" },
  //     { Name: "CasualLeave", Type: "Number" },
  //     { Name: "CasualLeaveUsed", Type: "Number" },
  //     { Name: "OtherLeaveUsed", Type: "Number" },
  //     { Name: "MaternityLeave", Type: "Number" },
  //     { Name: "MaternityLeaveUsed", Type: "Number" },
  //     { Name: "PaternityLeave", Type: "Number" },
  //     { Name: "PaternityLeaveUsed", Type: "Number" },
  //     { Name: "SickLeaveBalance", Type: "Calculated", Formula: "=(SickLeave-SickLeaveUsed)" },
  //     { Name: "OtherLeaveBalance", Type: "Calculated", Formula: "=(OtherLeave-OtherLeaveUsed)" },
  //     { Name: "PaternityLeaveBalance", Type: "Calculated", Formula: "=(PaternityLeave-PaternityLeaveUsed)" },
  //     { Name: "MaternityLeaveBalance", Type: "Calculated", Formula: "=(MaternityLeave-MaternityLeaveUsed)" },
  //     { Name: "CasualLeaveBalance", Type: "Calculated", Formula: "=(CasualLeave-CasualLeaveUsed)" },
  //     { Name: "EarnedLeave", Type: "Number" },
  //     { Name: "EarnedLeaveUsed", Type: "Number" },
  //     { Name: "EarnedLeaveBalance", Type: "Calculated", Formula: "=(EarnedLeave-EarnedLeaveUsed)" },
  //     { Name: "EmployeeID", Type: "SingleLine" },
  //     { Name: "StartDate", Type: "SingleLine" },
  //     { Name: "EndDate", Type: "SingleLine" },
  //     { Name: "Manager", Type: "Person" },



  //   ]
  //   await NewWeb.lists.add(ListName).then(() => {
  //     Columns.map(async (item: any) => {
  //       if (item.Type == "SingleLine") {
  //         NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addText(item.Name, 255, {
  //           Group: "Custom Column",
  //         }).then(() => {
  //           NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
  //           console.log(`${item.Name} column created successfully`)
  //           const progress = (1 * 100 / 75);
  //           this.updateProgress(progress);
  //         })
  //       }
  //       else if (item.Type == "MultiLine") {
  //         NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addMultilineText(item.Name, 255, true, false, false, false, {
  //           Group: "Custom Column",
  //           RichText: false,
  //         }).then(() => {
  //           NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
  //           console.log(`${item.Name} column created successfully`)
  //           const progress = (1 * 100 / 75);
  //           this.updateProgress(progress);
  //         })
  //       }
  //       else if (item.Type == "Number") {
  //         NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addNumber(item.Name).then(() => {
  //           NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
  //           console.log(`${item.Name} column created successfully`)
  //           const progress = (1 * 100 / 75);
  //           this.updateProgress(progress);
  //         })
  //       }
  //       else if (item.Type == "Calculated") {
  //         NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addCalculated(item.Name)
  //           .then(async () => {
  //             const progress = (1 * 100 / 75);
  //             this.updateProgress(progress);
  //             // NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
  //             // NewWeb.lists.getByTitle(ListName).fields.getByTitle(item.Name).update({ Formula: item.Formula },);
  //             await Promise.all([
  //               NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name),
  //               NewWeb.lists.getByTitle(ListName).fields.getByTitle(item.Name).update({ Formula: item.Formula })
  //             ]);
  //           }).then(() => {
  //             console.log(`${item.Name} column created successfully`)
  //           })
  //       }
  //       else if (item.Type == "Person") {
  //         NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addUser(item.Name).then(() => {
  //           NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
  //           console.log(`${item.Name} column created successfully`)
  //           const progress = (1 * 100 / 75);
  //           this.updateProgress(progress);
  //         })
  //       }
  //     })
  //     // Execute the batch
  //     batch.execute().then(function () {
  //       console.log("Batch operations completed successfully for creating " + ListName + " list");
  //     }).catch(function (error: any) {
  //       console.log("Error in batch operations for creating " + ListName + " list: " + error);
  //     });
  //   })
  // }
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
            const progress = (1 * 100 / 75);
            this.updateProgress(progress);
          })
        }
        else if (item.Type == "Date") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addDateTime(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 75);
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
            const progress = (1 * 100 / 75);
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
            const progress = (1 * 100 / 75);
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
  public async createApproverConfigurationList() {
    var ListName = "Approver Configuration";
    var batch = NewWeb.createBatch();
    var Columns = [
      { Name: "Approver", Type: "Person" },

    ]
    await NewWeb.lists.add(ListName).then(() => {
      Columns.map(async (item: any) => {
        if (item.Type == "Person") {
          NewWeb.lists.getByTitle(ListName).fields.inBatch(batch).addUser(item.Name).then(() => {
            NewWeb.lists.getByTitle(ListName).defaultView.fields.add(item.Name)
            console.log(`${item.Name} column created successfully`)
            const progress = (1 * 100 / 75);
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
    NewWeb.lists.getByTitle(ListName).items.add({
      ApproverId: this.state.CurrentUserId
    })
  }
  public updateProgress(value: any) {
    overAllValue += value;
    console.log("Progress", overAllValue)
    var RoundedValue = Math.ceil(overAllValue);
    if (RoundedValue >= progressEndValue) {
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
  public async createBalanceCollectionList() {
    const ListName = "BalanceCollection";
    const Columns = [
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
      // { Name: "Manager", Type: "Person" }
    ];

    try {
      // Add the list
      await NewWeb.lists.add(ListName);

      // Add columns
      for (const item of Columns) {
        if (item.Type === "Calculated") {
          await this.addCalculatedFieldWithRetry(ListName, item.Name, item.Formula);
        } else {
          await this.addField(ListName, item.Name, item.Type);
        }
      }

      console.log("Batch operations completed successfully for creating " + ListName + " list");
    } catch (error) {
      console.error("Error creating BalanceCollections list:", error);
    }
  }
  public async addField(listName: any, fieldName: any, fieldType: any) {
    try {
      let field;
      if (fieldType === "SingleLine") {
        field = await NewWeb.lists.getByTitle(listName).fields.addText(fieldName, 255, { Group: "Custom Column" });
        const progress = (1 * 100 / 75);
        this.updateProgress(progress);
      } else if (fieldType === "MultiLine") {
        field = await NewWeb.lists.getByTitle(listName).fields.addMultilineText(fieldName, 255, true, false, false, false, { Group: "Custom Column", RichText: false });
        const progress = (1 * 100 / 75);
        this.updateProgress(progress);
      } else if (fieldType === "Number") {
        field = await NewWeb.lists.getByTitle(listName).fields.addNumber(fieldName);
        const progress = (1 * 100 / 75);
        this.updateProgress(progress);
      } else if (fieldType === "Person") {
        field = await NewWeb.lists.getByTitle(listName).fields.addUser(fieldName);
        const progress = (1 * 100 / 75);
        this.updateProgress(progress);
      }

      if (field) {
        await NewWeb.lists.getByTitle(listName).defaultView.fields.add(fieldName);
        console.log(`${fieldName} column created successfully`);
      }
    } catch (error) {
      console.error(`Error adding ${fieldType} field ${fieldName}:`, error);
    }
  }
  public async addCalculatedFieldWithRetry(listName: any, fieldName: any, formula: any, retries = 0) {
    try {
      await this.addCalculatedField(listName, fieldName, formula);
      console.log(`${fieldName} field added successfully`);
    } catch (error) {
      if (error.statusCode === 409 && retries < MAX_RETRIES) {
        console.log(`Conflict detected, retrying (${retries + 1}/${MAX_RETRIES})...`);
        await new Promise(resolve => setTimeout(resolve, RETRY_DELAY_MS));
        await this.addCalculatedFieldWithRetry(listName, fieldName, formula, retries + 1);
      } else {
        console.error(`Error adding calculated field ${fieldName}:`, error);
      }
    }
  }
  public async addCalculatedField(listName: any, fieldName: any, formula: any) {
    try {
      // Add the calculated field
      await NewWeb.lists.getByTitle(listName).fields.addCalculated(fieldName, formula);
      const progress = (1 * 100 / 75);
      this.updateProgress(progress);
      // Update additional properties
      const field = NewWeb.lists.getByTitle(listName).fields.getByTitle(fieldName);
      await field.update({
        FieldTypeKind: FieldTypes.Calculated,
        Group: "Custom Column"
      });

      console.log(`${fieldName} field added successfully`);
    } catch (error) {
      console.error("Error adding calculated field:", error);
      throw error;
    }
  }
  public async addCurrentUserDetails() {
    await NewWeb.lists.getByTitle("BalanceCollection").items.add({
      EmployeeEmail: this.state.Empemail,
      SickLeave: 12,
      SickLeaveUsed: 0,
      OtherLeave: 100,
      OtherLeaveUsed: 0,
      EmployeeName: this.state.CurrentUserName,
      CasualLeave: 12,
      CasualLeaveUsed: 0,
      MaternityLeave: 130,
      MaternityLeaveUsed: 0,
      PaternityLeave: 40,
      PaternityLeaveUsed: 0,
      EarnedLeave: 12,
      EarnedLeaveUsed: 0,

    })
  }
  /** Check if the current user is in Owners group **/
  public async isOwnerGroupMember() {
    var reacthandler = this;
    let userDetails = await this.spLoggedInUser(this.props.context);

    let userID = userDetails.Id;
    console.log(userID);
    $.ajax({

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
  public GetUserlistitems() {
    var reactHandler = this;
    NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "*", "StartDate", "EndDate", "Reason", "Days", "Requester", "EmployeeEmail", "Day", "LeaveType", "Status", "AppliedDate", "CompOff").filter(`Author/Id eq ${this.props.userId}`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()


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
  private async spLoggedInUser(ctx: any) {
    try {
      const web = Web(ctx.pageContext.site.absoluteUrl);
      return await web.currentUser.get();
    } catch (error) {
      console.log("Error in spLoggedInUserDetails : " + error);
    }
  }
  public loadTable() {


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
          }).then(async () => {
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
            this.EmailSend(items)
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
  public Update_Blance_Count(result: any[], totaldaysapplied_leave: number, leavetype: string, LeaveStatus: any) {

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
                this.EmailSend(items)
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
            this.EmailSend(items)

          })
          this.Get_Blance_Count(totalDays, LeaveType, LeaveStatus)
        }

      }
    })


  }
  public async EmailSend(items: any) {
    const emailProps: IEmailProperties = {
      To: [items.EmployeeEmail, items.ApproverEmail], // Add the additional email address here
      Subject: 'Leave Request is Cancelled by ' + this.state.CurrentUserName,
      Body: `Leave Request Details<br/><br/>
              Status                    : Cancelled<br/><br/>
              Approver Name             : ${items.Approver}<br/><br/>
              Leave Type                : ${items.LeaveType}<br/><br/>
              Half Day / Full Day       : ${items.Day}<br/><br/>
              Start Date                : ${items.StartDate}<br/><br/>
              End Date                  : ${items.EndDate}<br/><br/>
              Compensation Date         : ${items.CompOff != null ? items.CompOff : "-"}<br/><br/>
              Reason                    : ${items.Reason}<br/><br/>
              Manager Comments (if any) : ${items.ManagerComments}<br/><br/>`,
      AdditionalHeaders: {
        "content-type": "text/html"
      }
    };

    await sp.utility.sendEmail(emailProps)
      .then((result) => {
        console.log(result)
      });
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
        var items = SpecificDate
        this.EmailSend(items)
        var ReduceLeaveDays = TotalDaysLeaveApplied - LeaveDates.length
        this.Get_Blance_Count(ReduceLeaveDays, LeaveTypee, LeaveStatuss)

      })

    } else {
      swal({
        text: "Don't select inbetween date",
        icon: "error"
      });
    }
  }
  public showLeaveMgmtDashboard() {
    this.setState({
      LeaveMgmtDashboard: true,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: false,
      PermissionApprovalDashboard: false,
      Approvals: false

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
      PermissionApprovalDashboard: false,
      Approvals: false

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
      PermissionApprovalDashboard: false,
      Approvals: false

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
      PermissionApprovalDashboard: false,
      Approvals: false

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
      PermissionApprovalDashboard: false,
      Approvals: false

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
      PermissionApprovalDashboard: false,
      Approvals: false

    })
  }
  public showApprovalsDashboard() {
    this.setState({
      LeaveMgmtDashboard: false,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: false,
      PermissionApprovalDashboard: false,
      Approvals: true
    })
  }
  public showPermissionApprovalsDashboard() {
    this.setState({
      LeaveMgmtDashboard: false,
      Holiday: false,
      LeaveMgmt: false,
      PermissionDashboard: false,
      PermissionRequest: false,
      AboutUs: false,
      Approvals: false,
      PermissionApprovalDashboard: true
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

        <span>{(item.OtherLeaveUsed)}/{parseFloat(item.OtherLeaveBalance)}</span>

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
                <img src={require("../img/logosmall.png")} alt="image" />
              </div>
              <div className="header-title"><h3>Leave Management System</h3></div>
              <div className="notification-part">
                <ul>
                  <li className="person-details">
                    <span id="CurrentUser-displayname">{this.state.CurrentUserName}</span>
                    <a href={`${this.props.siteurl}/_layouts/SignOut.aspx`} onClick={this.logout}><img src={require("../img/logout.png")} /></a>
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
              {this.state.IsCurrentUserisManager == true &&
                <>
                  <li className={this.state.Approvals == true ? "active" : ""} onClick={() => this.showApprovalsDashboard()}>Leave Approvals  </li>
                  <li className={this.state.PermissionApprovalDashboard == true ? "active" : ""} onClick={() => this.showPermissionApprovalsDashboard()}> Permission Approvals  </li>
                </>
              }
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


                  </ul>
                  {this.state.IsAdmin == true &&

                    <a href="${this.props.siteurl}/Lists/LeaveRequest/Approvedlist.aspx" className="btn btn-outline leave-req-link " id="submit">View leave list</a>
                  }
                  <button className="btn btn-outline" id="submit" onClick={() => this.showLeaveMgmt()}> New Leave Request  </button>

                </div>

                <div className="tab-content">
                  <div id="home" className="tab-pane fade in active">
                    <div className="three-blocks-wrap">

                      <div className="row">
                        <div className="col-md-4 leavecount-box" onClick={() => { this.setState({ leaveType: "TMX001" }); this.showLeaveMgmt() }}>
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/approved.png")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Earned Leave </p>

                              {EarnedLeaveBodycontent}
                            </div>

                          </div>
                        </div>
                        <div className="col-md-4 leavecount-box" onClick={() => { this.setState({ leaveType: "TMX002" }); this.showLeaveMgmt() }}>
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/pending.png")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Casual Leave </p>

                              {CasualLeaveBodycontent}

                            </div>

                          </div>
                        </div>
                        <div className="col-md-4 leavecount-box" onClick={() => { this.setState({ leaveType: "TMX003" }); this.showLeaveMgmt() }}>
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/sickleave.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Sick Leave </p>
                              {SickLeaveBodycontent}
                            </div>
                          </div>
                        </div>
                        <div className="col-md-4 leavecount-box" onClick={() => { this.setState({ leaveType: "TMX004" }); this.showLeaveMgmt() }}>
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/maternity.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Maternity Leave </p>
                              {MaternityLeaveBodycontent}
                            </div>
                          </div>
                        </div>
                        <div className="col-md-4 leavecount-box" onClick={() => { this.setState({ leaveType: "TMX005" }); this.showLeaveMgmt() }}>
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/paternityleave.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Paternity Leave </p>
                              {PaternityLeaveBodycontent}
                            </div>
                          </div>
                        </div>
                        <div className="col-md-4 leavecount-box" onClick={() => { this.setState({ leaveType: "TMX006" }); this.showLeaveMgmt() }}>
                          <div className="three-blocks">
                            <div className="three-blocks-img">
                              <img src={require("../img/otherleave.svg")} alt="image" />
                            </div>
                            <div className="three-blocks-desc">

                              <p> Unpaid Leave </p>
                              {OtherLeaveBodycontent}
                            </div>
                          </div>
                        </div>
                      </div>

                    </div>
                    <div className="table-wrap">
                      <div className="table-search-wrap clearfix">
                        <div className="table-search relative">

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


              </div>
            </div>
          </div>
        }
        {this.state.Holiday == true &&
          <Holiday description={''} leaveType={this.state.leaveType} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />
        }
        {this.state.LeaveMgmt == true &&
          <LeaveMgmt description={''} leaveType={this.state.leaveType} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.PermissionDashboard == true &&
          <PermissionDashboard description={''} leaveType={this.state.leaveType} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.PermissionRequest == true &&
          <PermissionRequest description={''} leaveType={this.state.leaveType} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.AboutUs == true &&
          <Aboutus description={''} leaveType={this.state.leaveType} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.Approvals == true &&
          <ApprovalDashboard description={''} leaveType={this.state.leaveType} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
        {this.state.PermissionApprovalDashboard == true &&
          <PermissionApprovalDashboard description={''} leaveType={this.state.leaveType} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
      </>
    );
  }

}

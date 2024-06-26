import * as React from 'react';
import { ILeaveMgmtDashboardProps } from './ILeaveMgmtDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp";
import { Web } from '@pnp/sp/webs';

import "datatables.net-dt/js/dataTables.dataTables";
import "datatables.net-dt/css/jquery.dataTables.min.css";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import * as $ from 'jquery';
import swal from "sweetalert";

import * as moment from 'moment';
import PermissionRequest from './PermissionRequest';
import "../css/style.css"
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";

let ItemId;
var CurrentUSERNAME = "";
var Usertype = "";
let NewWeb: any;

export interface PermissionDashboardState {
  DatatableItems: any[];
  Loggedinuserid: number;
  IsAdmin: boolean;
  CurrentUserName: string;
  CurrentUserDesignation: string;
  CurrentUserProfilePic: string;
  CurrentUserId: number;
  IsUser: boolean;
  Empemail: string;
  PermissionDashboard: boolean;
  PermissionRequest: boolean;
}

export default class PermissionDashboard extends React.Component<ILeaveMgmtDashboardProps, PermissionDashboardState> {

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
      Loggedinuserid: null,
      IsAdmin: false,
      CurrentUserName: "",
      CurrentUserDesignation: "",
      CurrentUserProfilePic: "",
      CurrentUserId: null,
      IsUser: false,
      Empemail: "",
      PermissionDashboard: true,
      PermissionRequest: false,
    };
    NewWeb = Web("" + this.props.siteurl + "")

  }

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

        reacthandler.setState({

          CurrentUserName: Name,

          CurrentUserDesignation: Designation,
          Empemail: email,

          CurrentUserProfilePic: `${reacthandler.props.siteurl}/_layouts/15/userphoto.aspx?size=l&username=${email}`

        });

      },

      error: function (jqXHR, textStatus, errorThrown) {

      }

    });

  }
  public GetUserlistitems() {
    var reactHandler = this;

    NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "EmployeeEmail", "Reason", "Status").filter(`Author/Id eq ${this.props.userId}`).orderBy("Created", false).top(5000).get()

      .then((items: any) => {
        if (items.length != 0) {

          reactHandler.setState({
            DatatableItems: items
          });
          this.loadTable();
          $("#no_data").hide()

        }
        else {
          this.loadTable();
          $("#no_data").show()
        }
      });


  }
  public GetAdminlistitems() {
    this.setState({ IsAdmin: true });
    var reactHandler = this;

    NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "EmployeeEmail", "Reason", "Status").orderBy("Created", false).top(5000).get()

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
  private async spLoggedInUser(ctx: any) {
    try {
      const web = Web(ctx.pageContext.site.absoluteUrl);
      return await web.currentUser.get();
    } catch (error) {
      console.log("Error in spLoggedInUserDetails : " + error);
    }
  }
  public async componentDidMount() {
    this.GetCurrentUserDetails();
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    let userDetails = await this.spLoggedInUser(this.props.context);
    console.log(userDetails.Id);
    let userID = userDetails.Id;
    this.setState({ CurrentUserId: userID });
    await this.isOwnerGroupMember();

    $(document).on('click', '#permission-dashboard', () => {
      this.setState({
        PermissionDashboard: true,
        PermissionRequest: false
      })
    })

  }
  public GetPermissionDetails() {
    var reactHandler = this;


    var url = "" + this.props.siteurl + "/_api/web/lists/getbytitle('EmployeePermission')/items?$select=PermissionHour,TimeUpto,PermissionOn,timefromwhen,Requester,Reason,Status&$orderby=Created desc";

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


      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }
  public async GetListitems() {
    var reactHandler = this;
    var UserType = "";
    let groups = await NewWeb.currentUser.groups();
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].Title == 'LMS Admin') {
        UserType = "Admin";

        console.log(UserType);
        Usertype = UserType;
        return Usertype;

      } else {

        UserType = "User";
        Usertype = "User"

      }
    }

    return Usertype;

  }
  public loadTable() {

    ($('#LMSDashboard') as any).DataTable({
      pageLength: 5,
      "bSort": false,
      "bDestroy": true,

      lengthMenu: [[5, 10, 20, 50, -1], [5, 10, 20, 50, "All"]],

      initComplete: function () {

        this.api().columns().every(function () {

          var column = this;
          var select = $('<select><option value="">All</option></select>')

            .appendTo($(column.header()).empty()).on('change', function () {

              var val = ($ as any).fn.dataTable.util.escapeRegex(

                ($(this) as any).val()

              );

              column.search(val ? '^' + val + '$' : '', true, false).draw();


            });

          column.data().unique().sort().each(function (d: string, j: any) {

            var temp2 = d;
            if (temp2.indexOf(">") != -1) {
              var temp = d.split(">");
              var temporary = temp[3].split("<");
              select.append('<option value="' + temporary[0] + '">' + temporary[0] + '</option>')
            } else {
              select.append('<option value="' + d + '">' + d + '</option>')
            }

          });


        });

      }

    });

  }
  public Cancel_Request_(itemidno: number) {

    swal({
      title: ` "Are you sure?"`,
      text: "Would you like to cancel the permission?",
      icon: "warning",
      buttons: ["No", "Yes"],
      dangerMode: true,
    } as any).then((willdelete) => {
      if (willdelete) {
        NewWeb.lists.getByTitle("EmployeePermission").items.getById(itemidno).update({
          Status: "Cancelled",
          CancelledBy: this.state.CurrentUserName
        }).then(() => {
          NewWeb.lists.getByTitle("EmployeePermission").items.select("*").filter(`ID eq ${itemidno}`).get()
            .then(async (items: any) => {
              const emailProps: IEmailProperties = {
                To: [items[0].EmployeeEmail, items[0].ApproverEmail], // Add the additional email address here
                Subject: 'Permission Request is Cancelled by ' + items[0].Approver + '',
                Body: `Permission Request Details<br/><br/>
                            Status                    : Cancelled<br/><br/>
                            Approver Name             : ${items[0].Approver}<br/><br/>
                            Permission On             : ${items[0].timefromwhen}<br/><br/>
                            Permission Hours          : ${items[0].PermissionHour}<br/><br/>
                            End Time                  : ${items[0].TimeUpto}<br/><br/>
                            Reason                    : ${items[0].Reason}<br/><br/>
                            Manager Comments (if any) : ${items[0].ManagerComments}<br/><br/>`,
                AdditionalHeaders: {
                  "content-type": "text/html"
                }
              };

              await sp.utility.sendEmail(emailProps)
                .then((result) => {
                  console.log(result)
                })
            });
          swal({
            text: "Permission cancelled successfully!",
            icon: "success",
          }).then(() => {
            location.reload()
          });

        })



      }
    })
  }
  public showPermissionRequest() {
    this.setState({
      PermissionDashboard: false,
      PermissionRequest: true,
    })
  }
  public render(): React.ReactElement<ILeaveMgmtDashboardProps> {
    let count = 0;
    let handler = this;

    const DataTableBodycontent: JSX.Element[] = this.state.DatatableItems.map(function (item, key) {

      count++;
      return (

        <tr id={`${key}-row-id`}>
          <td>{key + 1}</td>
          {handler.state.IsAdmin == true &&
            <td>{item.Requester}</td>
          }
          <td>{moment(item.PermissionOn).format('DD-MM-YYYY')}</td>
          <td>{item.timefromwhen}</td>
          <td>{item.TimeUpto}</td>
          <td>{item.PermissionHour}</td>
          <td className="reason-td">{item.Reason}</td>

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
          <td>
            {(item.State !== "Cancel" || item.State !== "Cancelled") &&
              <>
                {(handler.state.Empemail == item.EmployeeEmail && item.Status != "Cancelled" && item.Status != "Rejected" && moment(item.timefromwhen, "DD-MM-YYYY hh:mm A").isSameOrAfter(moment(), 'day')) &&


                  <p onClick={() => handler.Cancel_Request_(item.Id)}><img src={require("../img/cancel.svg")} alt="image" /></p>
                }

              </>
            }
          </td>

        </tr>

      );
    });

    return (
      <>
        {this.state.PermissionDashboard == true &&

          <div>
            <div className="container">
              <div className="dashboard-wrap">

                <div className="tab-headings clearfix">
                  <ul className="nav nav-pills">


                  </ul>

                  {this.state.IsAdmin == true &&
                    <a href={`${this.props.siteurl}/Lists/EmployeePermission/AllItems.aspx`} className="btn btn-outline leave-req-link" target='_blank' id="submit">View permission list</a>}
                  <button className="btn btn-outline" id="submit" onClick={() => this.showPermissionRequest()}> Permission Request  </button>
                </div>

                <div className="tab-content">
                  <div id="home" className="tab-pane fade in active">

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
                            <th className="text-center"> Status  </th>
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
                            <th>Start Time</th>
                            <th>End Time</th>
                            <th>Permission Hours</th>
                            <th className="reason-td">Reason</th>
                            <th className="text-center"> Status  </th>
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
        {this.state.PermissionRequest == true &&
          <PermissionRequest description={''} leaveType={''} context={this.props.context} siteurl={this.props.siteurl} userId={this.props.userId} />

        }
      </>
    );
  }
}

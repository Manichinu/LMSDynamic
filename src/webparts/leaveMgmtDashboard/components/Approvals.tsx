import * as React from 'react';
import styles from './LeaveMgmtDashboard.module.scss';
import { ILeaveMgmtDashboardProps } from './ILeaveMgmtDashboardProps';
// import { IPermissionDashboardState } from './IPermissionDashboardState';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp";
import { Web } from '@pnp/sp/webs';

import "datatables.net-dt/js/dataTables.dataTables";
import "datatables.net-dt/css/jquery.dataTables.min.css";
//import "datatables.net-dt/js/dataTables.dataTables";
//import "datatables.net-dt/css/jquery.dataTables.min.css";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import * as $ from 'jquery';
import swal from "sweetalert";

import * as moment from 'moment';
import PermissionRequest from './PermissionRequest';
import "../css/style.css"
import Swal from "sweetalert2";

let ItemId;
var CurrentUSERNAME = "";
var Usertype = "";
// const NewWeb = Web('https://tmxin.sharepoint.com/sites/ER/');
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

export default class ApprovalDashboard extends React.Component<ILeaveMgmtDashboardProps, PermissionDashboardState> {

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

            // url: `${reacthandler.props.siteurl}/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq  + ${this.props.userId}`,
            url: `${reacthandler.props.siteurl}/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq ${userID}`,

            type: "GET",

            headers: { 'Accept': 'application/json; odata=verbose;' },

            success: function (resultData) {

                if (resultData.d.results.length == 0) {
                    console.log("User not in group : LMS Admin Owners");



                } else {

                    console.log("User in group : LMS Admin Owners");


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
                reacthandler.GetUserlistitems()

            },

            error: function (jqXHR, textStatus, errorThrown) {

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
        //await this.GetListitems();
        // console.log("User Type:"+Usertype);
        // this.Checkusertype(Usertype);


        //  this.loadTable();
        $(document).on('click', '#permission-dashboard', () => {
            this.setState({
                PermissionDashboard: true,
                PermissionRequest: false
            })
        })

    }
    public GetUserlistitems() {
        var reactHandler = this;
        NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "*", "StartDate", "EndDate", "Reason", "Days", "Requester", "EmployeeEmail", "Day", "LeaveType", "Status", "AppliedDate", "CompOff").filter(`ApproverEmail eq '${this.state.Empemail}'`).expand('AttachmentFiles').orderBy("Created", false).top(5000).get()

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
    public loadTable() {
        //($('#LMSDashboard') as any).DataTable.destroy();

        ($('#LMSDashboard') as any).DataTable({
            pageLength: 5,
            "bSort": false,
            "bDestroy": true,

            lengthMenu: [[5, 10, 20, 50, -1], [5, 10, 20, 50, "All"]],

            initComplete: function () {

                this.api().columns().every(function () {

                    var column = this;
                    //  var select = $('<select class="form-control"><option value="">All</option></select>')
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
        //  }, 500);

    }
    public Approve(id: any) {
        Swal.fire({
            title: "<p>Comments</p>",
            html: "<textarea id='comments' /></textarea>",
            confirmButtonText: "Submit",
            customClass: {
                container: 'cancel-date',
            },
            showCloseButton: true,
            allowOutsideClick: true,
        }).then((result) => {
            if (result.isConfirmed) {
                NewWeb.lists.getByTitle("LeaveRequest").items.getById(id).update({
                    Status: "Approved",
                    ManagerComments: $("#comments").val()
                }).then(() => {
                    swal({
                        text: "Approved successfully!",
                        icon: "success",
                    }).then(() => {
                        location.reload()
                    });
                })
            }
        });
    }
    public Reject(id: any) {
        Swal.fire({
            title: "<p>Comments</p>",
            html: "<textarea id='comments' /></textarea>",
            confirmButtonText: "Submit",
            customClass: {
                container: 'cancel-date',
            },
            showCloseButton: true,
            allowOutsideClick: true,
            preConfirm: () => {
                var Comments = $("#comments").val();
                if (Comments == "") {
                    Swal.showValidationMessage("Please enter a comment");
                }
                return Comments;
            },
        }).then((result) => {
            if (result.isConfirmed) {
                NewWeb.lists.getByTitle("LeaveRequest").items.getById(id).update({
                    Status: "Rejected",
                    ManagerComments: $("#comments").val()
                }).then(() => {
                    swal({
                        text: "Rejected successfully!",
                        icon: "success",
                    }).then(() => {
                        location.reload()
                    });
                })
            }
        });
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
                                {(item.Status == "Pending" && moment(item.EndDate, "YYYY-MM-DD").isAfter(moment(), 'day')) &&
                                    <>
                                        <button onClick={() => handler.Approve(item.ID)}>Approve</button>
                                        <button onClick={() => handler.Reject(item.ID)}>Reject</button>
                                    </>
                                }

                            </>
                        }
                    </td>
                </tr>

            );
        });

        return (
            <>

                <div className={styles.permissionDashboard} >
                    <div className="container">
                        <div className="dashboard-wrap">

                            <div className="tab-headings clearfix">
                                <ul className="nav nav-pills">


                                </ul>

                            </div>

                            <div className="tab-content">
                                <div id="home" className="tab-pane fade in active">

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
                        </div>
                    </div>
                </div>

            </>
        );
    }
}
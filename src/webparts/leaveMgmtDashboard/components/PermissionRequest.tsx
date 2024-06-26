import * as React from 'react';
import { ILeaveMgmtDashboardProps } from './ILeaveMgmtDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "react-datepicker/dist/react-datepicker.css";
import 'bootstrap/dist/css/bootstrap.min.css';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import * as moment from "moment-timezone";

import DatePicker from 'react-datepicker';
import swal from "sweetalert";
import * as $ from "jquery";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { IAttachmentFileInfo, IItemAddResult, Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "../css/style.css";
import { sp } from "@pnp/sp";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";

let datesCollection: string[] = [];
var PreviousLeaveRequestDates: any[] = [];
var PreviousPermissionRequestDates = [];
let IsValidReqularRequest = false;
var Approver_Manager_Details: any = []
let NewWeb: any;

export interface IPermissionRequestState {
  startDate: any;
  selectedtime: string;
  CurrentUserName: string;
  CurrentUserDesignation: string;
  CurrentUserProfilePic: string;
  Email: string;
  IsAlreadyexist: boolean;
  CurrentUserId: number;
  Appliedleaveitems: any[];

}
export default class PermissionRequest extends React.Component<ILeaveMgmtDashboardProps, IPermissionRequestState> {

  public constructor(props: ILeaveMgmtDashboardProps, state: IPermissionRequestState) {

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


    this.state = {

      startDate: new Date(),
      selectedtime: "",
      CurrentUserName: "",
      CurrentUserDesignation: "",
      CurrentUserProfilePic: "",
      Email: "",
      IsAlreadyexist: false,
      CurrentUserId: null,
      Appliedleaveitems: [],

    };
    NewWeb = Web("" + this.props.siteurl + "")

    this.handleChange = this.handleChange.bind(this);
    this.handleSelect = this.handleSelect.bind(this);

  }
  handleSelect(date: moment.MomentInput) {
    var selectedhr: any = $('#ddl-Permissionhr').val();
    var finaltime = moment(date, "DD-MM-YYYY hh:mm A").add(selectedhr, 'hours').format('D-MM-YYYY hh:mm A');

    $("#txt-EndDate").val(finaltime);
    this.setState({ startDate: date });
  }
  handleChange(date: moment.MomentInput) {
    var selectedhr: any = $('#ddl-Permissionhr').val();
    var finaltime = moment(date, "DD-MM-YYYY hh:mm A").add(selectedhr, 'hours').format('D-MM-YYYY hh:mm A');

    $("#txt-EndDate").val(finaltime);
    this.setState({ startDate: date });
  }
  public componentDidMount() {
    this.GetCurrentUserDetails();
  }
  public isInArray(PreviousLeaveRequestDates: any, value: string) {
    return (PreviousLeaveRequestDates.find((item: any) => { return item == value }) || []).length > 0;
  }
  public getDaysBetweenDates(startDate: moment.Moment, endDate: moment.Moment) {
    var now = startDate.clone();
    while (now.isSameOrBefore(endDate)) {
      PreviousLeaveRequestDates.push(now.format('YYYY-MM-DD'));
      now.add(1, 'days').format('DD/MM/YYYY');
    }
    return PreviousLeaveRequestDates;
  };
  public GetPreviousLeaveRequestDates(email: any) {


    var filterquery = `EmployeeEmail eq '${email}' and Status ne 'Rejected'`// and enddate ge '${moment().format("DD-MM-YYYY")}'`
    NewWeb.lists.getByTitle("LeaveRequest").items.select("StartDate", "EndDate", "EmployeeEmail").filter(filterquery).orderBy("Created", false).get().then((response: any): void => {
      if (response.length != 0) {
        let i;
        for (i = 0; i < response.length;) {
          var From = response[i].StartDate;
          console.log(From);

          var To = response[i].EndDate;
          console.log(To);

          var tempFromDate = moment(From).format("YYYY-MM-DD");
          console.log(tempFromDate);
          var tempToDate = moment(To).format("YYYY-MM-DD");
          console.log(tempToDate);
          var dateList = this.getDaysBetweenDates(moment(tempFromDate), moment(tempToDate));
          console.log("dateList LeaveRequest: " + dateList);
          i++;
        }
      }
    });
  }
  public GetPreviousPermissionRequestDates(email: any) {

    var filterquery = `EmployeeEmail eq '${email}'and Status ne 'Rejected'`
    NewWeb.lists.getByTitle("EmployeePermission").items.select("timefromwhen", "EmployeeEmail").filter(filterquery).orderBy("Created", false).get().then((response: any): void => {
      if (response.length != 0) {
        let i;
        for (i = 0; i < response.length;) {
          var From = response[i].timefromwhen;

          var tempFromDate = moment(From, "DD-MM-YYYYTHH:mm").format('DD-MM-YYYY');
          var tempToDate = moment(From, "DD-MM-YYYYTHH:mm").format("DD-MM-YYYY");


          var dateList = this.getDaysBetweenDates(moment(tempFromDate), moment(tempToDate));
          console.log("dateList PermissionRequest: " + dateList);
          i++;
        }
      }
    });
  }
  public clearerror() {
    $("#divErrorText").empty();
    $("#divErrorText").hide();
  }
  public AlreadyexistinPR(email: any) {
    datesCollection = [];
    var reactHandler = this;
    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('EmployeePermission')/items?$select=timefromwhen,EmployeeEmail&$filter('Author/EmployeeEmail eq '${email}'')`;

    $.ajax({
      url: url,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {


        reactHandler.setState({
          Appliedleaveitems: resultData.d.results
        });

        for (var i = 0; i < resultData.d.results.length; i++) {
          var permdate = resultData.d.results[i].timefromwhen;
          var permconv = moment(permdate, "DD-MM-YYYYh:mm A").format('DD-MM-YYYY');
          datesCollection.push(permconv);
        }
        console.log(datesCollection);


      },

      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }
  public Calculatehours() {
    this.clearerror();
    var selectedhr: any = $('#ddl-Permissionhr').val();

    var selectedtime = this.state.startDate;

    var calculatedtime = moment(selectedtime, "YYYY-MM-DDTHH:mm").add(selectedhr, 'hours').format('D-MM-YYYY hh:mm A');

    $("#txt-EndDate").val(calculatedtime);

  }
  public getselectedstarttime(date: any) {

    this.setState({ selectedtime: date });
    var selectedhr: any = $('#ddl-Permissionhr').val();
    var selectedtime = this.state.selectedtime;
    console.log(selectedtime);
    var calculatedtime = moment(selectedtime, "YYYY-MM-DDTHH:mm").add(selectedhr, 'hours').format('D-MM-YYYY hh:mm A');
    $("#txt-EndDate").val(calculatedtime);

  }  
  public Checkalreadyinleave() {
    let Status = true;
    var selectedtime = this.state.startDate;

    var permissionDate = moment(selectedtime, "YYYY-MM-DDTHH:mm").format('YYYY-MM-DD');

    var filterquery = `EmployeeEmail eq '${this.state.Email}' and '${permissionDate}' ge startdate and '${permissionDate}' le enddate`
    debugger;

    NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate").filter(filterquery).orderBy("Created", false).get().then((response: any): void => {
      if (response.length != 0) {
        console.log(response);
        this.setState({
          IsAlreadyexist: true
        });
        Status = false;
      } else {
        var filterquery1 = `EmployeeEmail eq '${this.state.Email}' and  timefromwhen eq '${permissionDate}'`
        NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "timefromwhen", "PermissionOn").filter(filterquery1).orderBy("Created", false).get().then((response: any): void => {
          if (response.length != 0) {
            this.setState({
              IsAlreadyexist: true
            });
            Status = false;
          }
          else {
            this.setState({

              IsAlreadyexist: false

            });
            Status = true;

          }
        });

      }

    })

    return Status;

  }
  public LeaveformValidation() {

    var Formstatus = false;
    var ErrorMsg = "";
    $("#divErrorText").empty();

    var permissionhour = $("#ddl-Permissionhr").val();
    var startdate = $("#txt-Startdate").val();
    var enddate = $("#txt-EndDate").val();
    var Reason = $("#txt-reason").val();


    if (permissionhour == "") {
      ErrorMsg = "Please Select Permission Hour";
      Formstatus = true;
    } else if (Formstatus == false && startdate == "") {
      ErrorMsg = "Please Select StartDate";
      Formstatus = true;


    } else if (Formstatus == false && Reason == "") {
      ErrorMsg = "Please Enter Reason";
      Formstatus = true;
    }

    if (Formstatus) {
      $("#divErrorText").append(ErrorMsg);
      $("#divErrorText").show();
      return false;
    } else {
      $("#divErrorText").empty();
      $("#divErrorText").hide();
      return true;
    }


  }
  public GetCurrentUserDetails() {

    var reacthandler = this;

    $.ajax({

      url: `${reacthandler.props.siteurl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,

      type: "GET",

      headers: { 'Accept': 'application/json; odata=verbose;' },

      success: function (resultData) {
        console.log(resultData)
        var email = resultData.d.Email;

        var Name = resultData.d.DisplayName;

        var Designation = resultData.d.Title;
        var gender = resultData.d.Streetaddress;

        reacthandler.setState({

          CurrentUserName: Name,

          CurrentUserDesignation: Designation,

          Email: email,
          CurrentUserId: resultData.d.Id,
          CurrentUserProfilePic: `${reacthandler.props.siteurl}/_layouts/15/userphoto.aspx?size=l&username=${email}`

        });
        reacthandler.Get_CorrespondingApprover(email)
        reacthandler.GetPreviousLeaveRequestDates(email);
        reacthandler.GetPreviousPermissionRequestDates(email);

      },

      error: function (jqXHR, textStatus, errorThrown) {

      }

    });

  }
  public Addtolist() {


    var permissionhour = $("#ddl-Permissionhr").val();
    var enddate = $("#txt-EndDate").val();
    var Reason = $("#txt-reason").val();
    var selectedtime = this.state.startDate;

    var startdate = moment(selectedtime, "YYYY-MM-DDTHH:mm").format('YYYY-MM-DD');
    var starttime = moment(selectedtime, "YYYY-MM-DDTHH:mm").format('DD-MM-YYYY hh:mm A');
    var EndTime = moment(enddate, "DD-MM-YYYY hh:mm A").format('DD-MM-YYYY hh:mm A');

    var now = new Date();
    var day = ("0" + now.getDate()).slice(-2);
    var month = ("0" + (now.getMonth() + 1)).slice(-2);
    var today = now.getFullYear() + "-" + (month) + "-" + (day);
    var reactHandler = this;
    var curentURL = $(location).attr('href');
    var decodedUrl = decodeURIComponent(curentURL);
    const url: any = new URL(decodedUrl);
    const ItemId = url.searchParams.get("ItemID");

    if (this.isInArray(PreviousLeaveRequestDates, startdate) == false) {//6 not found

      if (this.LeaveformValidation()) {

        NewWeb.lists.getByTitle("EmployeePermission").items.add({
          PermissionHour: permissionhour,
          timefromwhen: starttime,
          TimeUpto: EndTime,
          Reason: Reason,
          PermissionOn: today,
          Requester: this.state.CurrentUserName,
          EmployeeEmail: this.state.Email,
          Approver: Approver_Manager_Details[0].ApproverName,
          ApproverEmail: Approver_Manager_Details[0].ApproverEmail,
          Status: "Pending"

        })

          .then((item: any) => {

            let ID = item.data.Id;
            NewWeb.lists.getByTitle("EmployeePermission").items.select("*").filter(`ID eq ${ID}`).get()
              .then(async (items: any) => {
                const emailProps: IEmailProperties = {
                  To: ['' + items[0].ApproverEmail + ''],
                  Subject: 'Permission Request is Raised by ' + this.state.CurrentUserName + '',
                  Body: `Permission Request Details<br/><br/>
                            Status                    : Pending<br/><br/>
                            Approver Name             : ${items[0].Approver}<br/><br/>
                            Permission On             : ${items[0].timefromwhen}<br/><br/>
                            Permission Hours          : ${items[0].PermissionHour}<br/><br/>
                            End Time                  : ${items[0].TimeUpto}<br/><br/>
                            Reason                    : ${items[0].Reason}<br/><br/>
                            <p>Please <a href='${this.props.siteurl}/SitePages/LeaveManagement.aspx?tab=permission'>click here</a> to view the request</p>`,
                  AdditionalHeaders: {
                    "content-type": "text/html"
                  }
                };

                await sp.utility.sendEmail(emailProps)
                  .then((result: any) => {
                    console.log(result)
                  })
              });

            swal({
              text: "Permission applied successfully!",
              icon: "success",

            }).then(() => {

              location.reload()

            });


          });


      }

    }
    else {

      swal({

        text: "Already leave request taken on selected date",
        icon: "error"
      });
    }
  }
  public Get_CorrespondingApprover(EmployeeEmailid: any) {
    var currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    NewWeb.lists.getByTitle("Approver Configuration").items.select("ID", "*", "Approver/Title", "Approver/EMail").expand("Approver").get()
      .then((result: any) => {
        if (result.length != 0) {
          console.log(result);

          Approver_Manager_Details.push({
            ApproverName: result[0].Approver.Title,
            ApproverEmail: result[0].Approver.EMail
          })

          console.log(Approver_Manager_Details)
        }
      })
  }
  public render(): React.ReactElement<ILeaveMgmtDashboardProps> {

    return (
      <div>

        <div className="container">
          <div className="dashboard-wrap">

            <div className="form-header">
              <a href=""><img src={require("../img/back.svg")} alt="image" /> </a> <span> Permission Request </span>
            </div>

            <div className="form-body">
              <div className="form-section">
                <div className="row">
                  <div className="col-md-4 col-sm-4 permission_date_picker">
                    <div className="form-group required relative">

                      <DatePicker
                        name="startDate"
                        selected={this.state.startDate}

                        onSelect={this.handleSelect}
                        onChange={(date) => this.handleChange(date)}//, 'startDate')} 
                        filterDate={(date) => date.getDay() != 6 && date.getDay() != 0}
                        showTimeSelect
                        timeFormat="HH:mm"
                        dateFormat="d-MM-yyyy h:mm aa"

                      />


                      <span className="floating-label ">Permission on</span>


                    </div>
                  </div>

                  <div className="col-md-4 col-sm-4">
                    <div className="form-group required relative">
                      <select name="Permission" id="ddl-Permissionhr" className="form-control" onChange={() => this.Calculatehours()}>
                        <option value="">--Select Hours--</option>
                        <option value="0.5">0.5 hour(s)</option>
                        <option value="1.0">1 hour(s)</option>
                        <option value="1.5">1.5 hour(s)</option>
                        <option value="2.0">2 hour(s)</option>
                        <option value="2.5">2.5 hour(s)</option>
                        <option value="3.0">3 hour(s)</option>
                        <option value="3.5">3.5 hour(s)</option>
                        <option value="4.0">4 hour(s)</option>
                        <option value="4.5">4.5 hour(s)</option>
                        <option value="5.0">5 hour(s)</option>
                        <option value="5.5">5.5 hour(s)</option>
                        <option value="6.0">6 hour(s)</option>
                        <option value="6.5">6.5 hour(s)</option>
                        <option value="7.0">7 hour(s)</option>
                        <option value="7.5">7.5 hour(s)</option>
                        <option value="8.0">8 hour(s)</option>
                      </select>
                      <span className="floating-label "> Permission Hours </span>
                    </div>
                  </div>
                  <div className="col-md-4 col-sm-4">
                    <div className="form-group required relative">
                      <input type="text" className="form-control read-only-class" id="txt-EndDate" readOnly />
                      <span className="floating-label "> End Time </span>
                    </div>
                  </div>
                </div>
                <div>
                  <div className="row">
                    <div className="col-md-8">
                      <div className="form-group required relative">
                        <div className="form-group">

                          <input type="text" className="form-control" id="txt-reason" maxLength={250} autoComplete="off" onKeyPress={() => this.clearerror()} />
                          <span className="floating-label ">Enter Reason</span>
                        </div>

                      </div>
                    </div>
                  </div>

                  <div className="row">
                    <div
                      className="alert alert-danger"
                      role="alert"
                      id="divErrorText"
                      style={{ display: "none" }}
                    ></div>
                    <div className="col-md-12 btn-padding">
                      <button className="btn btn-primary" id="submit" onClick={() => this.Addtolist()}>Submit</button>

                    </div>
                  </div>
                </div>
              </div>
            </div>

          </div>
        </div>
      </div >
    );
  }
}

import * as React from 'react';
import styles from './LeaveMgmtDashboard.module.scss';
import { ILeaveMgmtDashboardProps } from './ILeaveMgmtDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
// import { IHolidayState } from './IHolidayState';
import { Web } from '@pnp/sp/webs';
import "@pnp/sp/site-users/web";
import * as $ from 'jquery';
import * as moment from "moment";
import { _SiteGroups } from '@pnp/sp/site-groups/types';
// const NewWeb = Web('https://tmxin.sharepoint.com/sites/ER/');
let ItemId;
let NewWeb: any;

export interface HolidayState {
  HolidayItems: any[];
  CurrentUserName: string;
  CurrentUserDesignation: string;
  CurrentUserProfilePic: string;
  IsAdmin: boolean;
  CurrentUserId: number;
}

export default class Holiday extends React.Component<ILeaveMgmtDashboardProps, HolidayState> {
  public constructor(props: ILeaveMgmtDashboardProps, state: HolidayState) {
    super(props);

    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
    );
    SPComponentLoader.loadCss(`https://fonts.googleapis.com`);
    SPComponentLoader.loadCss(`https://fonts.gstatic.com" crossorigin`);
    SPComponentLoader.loadCss(`https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap`);
    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
    );
    SPComponentLoader.loadScript(
      `https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js`
    );
    SPComponentLoader.loadScript(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.js`
    );
    SPComponentLoader.loadCss(
      `${this.props.siteurl}/SiteAssets/LeavePortal/css/style.css?v=1.14`
    );

    this.state = {
      HolidayItems: [],
      CurrentUserId: null,
      IsAdmin: false,
      CurrentUserName: "",
      CurrentUserDesignation: "",
      CurrentUserProfilePic: ""
    };
    NewWeb = Web("" + this.props.siteurl + "")

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

          CurrentUserProfilePic: `${reacthandler.props.siteurl}/_layouts/15/userphoto.aspx?size=l&username=${email}`

        });

      },

      error: function (jqXHR, textStatus, errorThrown) {

      }

    });

  }
  public logout() {

    localStorage.clear();
    // window.location.href=` https://tmxin.sharepoint.com/sites/POC/SPIP/_layouts/closeConnection.aspx?loginasanotheruser=true`;
    window.location.href = `https://login.windows.net/common/oauth2/logout`;

  }
  public async componentDidMount() {
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    this.GetHolidaylist();
    this._spLoggedInUserDetails();
    this.GetCurrentUserDetails();

    await this.CheckManagerPermissionPrivillages();

  }
  public _spLoggedInUserDetails() {
    NewWeb.currentUser.get().then((user: any) => {
      let userID = user.Id;
      this.setState({ CurrentUserId: userID });
    }, (errorResponse: any) => {
      //console.log(errorResponse);
    }
    );
  }
  public GetHolidaylist() {
    var reactHandler = this;
    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('HolidayCollection')/items?$select=StartDate,HolidayName&$orderby=StartDate asc`;
    {/* NewWeb.lists.getByTitle("HolidayCollection").items.select("StartDate","HolidayName").top(10).orderBy("StartDate",true).get()
    .then((items)=>{
      if(items.length != 0){
        this.setState({
          HolidayItems:items
        });
      }
    });
  */}


    $.ajax({
      url: url,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        console.log(resultData);

        reactHandler.setState({
          HolidayItems: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }

  public async CheckManagerPermissionPrivillages() {

    /*  let SiteGroups = 'LMS Admin';
      let InGroup: boolean = false;
   
      let grp = await NewWeb.currentUser.groups.get().then((r: any) => {
        r.forEach((grp: _SiteGroups) => {
          if (grp["Title"] == SiteGroups) {
            InGroup = true;*/
    let groups = await NewWeb.currentUser.groups();
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].Title == 'LMS Admin') {

        this.setState({ IsAdmin: true }); //To enable admin access to Specific Group Users alone      

        break;

      } else {
        this.setState({ IsAdmin: false });


      }
      {/*   if(groups[i].Title != 'LMS Admin'){  
        this.setState({IsUser:true}); //To enable Manager access to Specific Group Users alone                             
        this.GetListitems("User");
        break;
      }else{
        this.setState({IsUser:false}); 
      }
    */}

    }
    // this.GetListitems("User");
  }
  public render(): React.ReactElement<ILeaveMgmtDashboardProps> {
    let count = 0;

    const HolidayBodycontent: JSX.Element[] = this.state.HolidayItems.map(function (item, key) {

      count++;
      return (
        <li>
          <div className="holiday-page-inner">
            <h4>{item.HolidayName} </h4>
            <p>{moment(item.StartDate).format('LL')}</p>

          </div>

        </li>
      );
    });

    return (
      <div className={styles.holiday}>        
        <div className="container">
          <div className="dashboard-wrap">
            {this.state.IsAdmin == true &&

              <a href={`${this.props.siteurl}/Lists/HolidayCollection/AllItems.aspx`} className="btn btn-outline leave-req-link " id="submit">View holiday list</a>}
            <div className="holiday-page">
              <p> Below is the list of our company’s paid holidays. Our offices will be closed for observance. </p>
              <ul>

                {HolidayBodycontent}


              </ul>
            </div>

          </div>
        </div>
      </div>
    );
  }
}

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
import "../css/style.css"

import * as moment from 'moment';
import PermissionRequest from './PermissionRequest';
let ItemId;
var CurrentUSERNAME = "";
var Usertype = "";
// const NewWeb = Web('https://tmxin.sharepoint.com/sites/ER/');
let NewWeb: any;

export interface AboutusState {
    DatatableItems: any[];
}

export default class Aboutus extends React.Component<ILeaveMgmtDashboardProps, AboutusState> {

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
        //     `${this.props.siteurl}/SiteAssets/LeavePortal/css/style.css?v=1.14`
        // );
        SPComponentLoader.loadScript(
            `https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js`
        );
        SPComponentLoader.loadScript(
            `https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.2/dist/umd/popper.min.js`
        );



        sp.setup({
            spfxContext: this.props.context
        });

        this.state = {
            DatatableItems: [],
        };
        NewWeb = Web("" + this.props.siteurl + "")

    }
    public componentDidMount(): void {
        this.getListItems();
    }
    public getListItems(): void {
        NewWeb.lists.getByTitle("LeaveTypeCollection").items.select("Types", "Details").get()
            .then((results: any) => {
                this.setState({
                    DatatableItems: results
                });
            })
            .catch((error: any) => {
                console.log("Failed to get list items!");
                console.log(error);
            });
    }



    public render(): React.ReactElement<ILeaveMgmtDashboardProps> {

        const LeaveTypes: any = this.state.DatatableItems.map((item: any, key: any) => {
            return (
                <>
                    <div className="accordion-item">
                        <h2 className="accordion-header" id={`headingOne${key}`}>
                            <button className="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target={`#collapse${key}`} aria-expanded="false" aria-controls={`collapse${key}`}>
                                {item.Types}
                            </button>
                        </h2>
                        <div id={`collapse${key}`} className="accordion-collapse collapse" aria-labelledby={`headingOne${key}`} data-bs-parent="#accordionExample">
                            <div className="accordion-body">
                                {item.Details}
                            </div>
                        </div>
                    </div>
                </>
            )
        })


        return (
            <>
                <div className="container">
                    <div className="dashboard-wrap">
                        <ul>
                            <li className="li-bold"> About Leave Types</li>
                        </ul>
                        <div className="accordion" id="accordionExample">
                            {LeaveTypes}                           
                        </div>
                    </div>
                </div>
            </>
        );
    }
}

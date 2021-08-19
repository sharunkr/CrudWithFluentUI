import * as React from "react";
import styles from "./Reactdemo.module.scss";
import { IReactdemoProps } from "./IReactdemoProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { useState } from "react";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


export default class PnpList extends React.Component<any, any> {
    constructor(props){
        super(props)
        sp.setup({spfxContext:this.props.context})
    }
    private addData = async () => {
        const data = await sp.web.lists.getByTitle("PnpList").items.add({
          Title: "Title1",
          name: "Äpple2",
          Uage: 3453,
          doj: new Date()
          
        });
        console.log(data);
      }
    private getData = async () => {
        const items: any[] = await sp.web.lists.getByTitle("PnpList").items.get();
        console.log(items);
      }
}
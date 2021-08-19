import * as React from "react";
import styles from "./Reactdemo.module.scss";
import { IReactdemoProps } from "./IReactdemoProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { useState } from "react";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { Calendar,defaultCalendarStrings } from "@fluentui/react";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  SelectionMode,
} from "office-ui-fabric-react/lib/DetailsList";


export interface IDetailsListBasicExampleItem {
  eid: number;
  name: string;
  age: number;
  empState: string;
  doj: number;
}

export interface IDetailsListBasicExampleState {
  items: IDetailsListBasicExampleItem[];
  selectionDetails: string;
}


//type MyState = { name, employees, age, empState, doj, };
export default class Reactdemo extends React.Component<any, any> {
  private empid = 0;
  private update = false;
  private Eid = 0;
  private selectedEmpData = [];
  private _columns: IColumn[];
  private _allItems: IDetailsListBasicExampleItem[];
  private _selection: Selection;
  private all = [];
  private selection = 0;
  constructor(props) {
    super(props);
    sp.setup({ spfxContext: this.props.context });
    this.state = {
      employees: [],
      storeListdata: [],
      All: [],
      Eid: null,
      name: "",
      age: null,
      empState: "",
      doj: null,
      items: [],
      selectionDetails: "",
      item: 0,
      disabled: true,
    };
    this.SaveData = this.SaveData.bind(this);
    this.UpdateData = this.UpdateData.bind(this);
    this.DeleteData = this.DeleteData.bind(this);

    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ selectionDetails: this._getSelectionDetails() }),
    });
    this._columns = [
      {
        key: "column1",
        name: "ID",
        fieldName: "eid",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column2",
        name: "NAME",
        fieldName: "name",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column3",
        name: "AGE",
        fieldName: "age",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column4",
        name: "EMPSTATE",
        fieldName: "empState",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column5",
        name: "DOJ",
        fieldName: "doj",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
    ];
  }
  componentDidMount() {
    this.getData();
  }

  getdata = (event) => {
    this.setState({
      [event.target.name]: event.target.value,
    });
  };

  SaveData = async (e) => {
    e.preventDefault();
    const iar: IItemAddResult = await sp.web.lists
      .getByTitle("PnpCrud")
      .items.add({
        Eid: this.state.Eid,
        name: this.state.name,
        age: this.state.age,
        empState: this.state.empState,
        doj: this.state.doj,
      });
    this.getData();
  };

  UpdateData = async (items) => {
    let list = sp.web.lists.getByTitle("PnpCrud");

    const i = await list.items.getById(items).update({
      Eid: this.state.Eid,
      name: this.state.name,
      age: this.state.age,
      empState: this.state.empState,
      doj: this.state.doj,
    });
    this.getData();
    this._selection.setAllSelected(false);
    this.handleDisable();
  };

  DeleteData = async (items) => {
    let list = sp.web.lists.getByTitle("PnpCrud");
    await list.items.getById(items).delete();
    this.getData();
    this._selection.setAllSelected(false);
    this.handleDisable();
  };
  private getData = async () => {
    await sp.web.lists
      .getByTitle("PnpCrud")
      .items.top(5000)
      .select("*")
      .orderBy("name", true)
      .get()
      .then((items) => {
        //this._allItems = items;
        let storeItems = [];
        for (let i = 0; i < items.length; i++) {
          storeItems.push({
            eid: items[i].Id,
            name: items[i].name,
            age: items[i].age,
            empState: items[i].empState,
            doj:
              new Date(items[i].doj).getDate() +
              "-" +
              (new Date(items[i].doj).getMonth() + 1) +
              "-" +
              new Date(items[i].doj).getFullYear(),
          });
        }

        this.setState({
          items: [...storeItems],
        });
      })
      .catch((e) => {
        console.log("error", e);
      });
  };

  private _getSelectionDetails() {
    const getitem =
      this._selection.getSelection()[0] as IDetailsListBasicExampleItem;
    const selectioncount = this._selection.getSelectedCount();
    this.selection = selectioncount;
    this.setState({
      item: getitem.eid,
    });
    this.handleEnable();
  }
  handleEnable = () => {
    if (this.selection > 0) {
      this.setState({
        disabled: false,
      });
    }
    this._selection.setAllSelected(false);
  };
  handleDisable = () => {
    this.setState({
      disabled: true,
    });
  };

  
  public render() {
    return (
      <div>
        <div>
          <h1>CRUD OPERATION</h1>
          <form
            onSubmit={
              !this.update
                ? (event) => this.SaveData(event)
                : (event) => this.UpdateData(event)
            }
          >
            <TextField
        
        placeholder="enter a number..."
              label="Eid"
              type="number"
              name="Eid"
              value={this.state.Eid}
              onChange={(event) => this.getdata(event)}
              required

            />
            <br />
            <TextField
            placeholder="enter a name..."
              label="Name"
              type="text"
              name="name"
              value={this.state.name}
              onChange={(event) => this.getdata(event)}
              required
            />
            <br />

            <TextField
            placeholder="enter your age..."
              label="Age"
              type="number"
              name="age"
              value={this.state.age}
              onChange={(event) => this.getdata(event)}
              required
            />
            <br />
            <TextField
            placeholder="enter your state..."
              label="EmpState"
              type="text"
              name="empState"
              value={this.state.empState}
              onChange={(event) => this.getdata(event)}
              required
            />
            <br />

            <TextField
            placeholder="select a date..."
              label="DOJ"
              type="date"
              name="doj"
              value={this.state.doj}
              onChange={(event) => this.getdata(event)}
              required
            />
            <br />
            <PrimaryButton type="submit">SAVE</PrimaryButton>
            <PrimaryButton
              onClick={() => this.UpdateData(this.state.item)}
              disabled={this.state.disabled}
            >
              EDIT
            </PrimaryButton>
            <DefaultButton
              onClick={() => this.DeleteData(this.state.item)}
              disabled={this.state.disabled}
            >
              DELETE
            </DefaultButton>
          </form>
          <div>
            <Fabric>
              <DetailsList
                items={this.state.items}
                columns={this._columns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selection={this._selection}
                selectionPreservedOnEmptyClick={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="Row checkbox"
                selectionMode={SelectionMode.single}
                // onItemInvoked={this._onItemInvoked}
              />
            </Fabric>
          </div>
        </div>
      </div>
    );
  }
}

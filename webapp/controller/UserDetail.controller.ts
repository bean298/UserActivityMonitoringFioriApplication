import MessageBox from "sap/m/MessageBox";
import Controller from "sap/fe/core/PageController";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import JSONModel from "sap/ui/model/json/JSONModel";
import ODataListBinding from "sap/ui/model/odata/v4/ODataListBinding";
import ODataModel from "sap/ui/model/odata/v4/ODataModel";
import Formatter from "useraudit/formatter/Formatter";
import Select from "sap/m/Select";
import Table from "sap/ui/table/Table";
import DatePicker from "sap/m/DatePicker";
import DateFormat from "sap/ui/core/format/DateFormat";
import Spreadsheet from "sap/ui/export/Spreadsheet";
import MessageToast from "sap/m/MessageToast";

export default class UserDetail extends Controller {
  public formatter = Formatter;

  /**
   * Called when the controller is initialized.
   **/
  public onInit(): void {
    super.onInit();

    const oRouter = (this as any).getAppComponent().getRouter();
    if (oRouter) {
      oRouter
        .getRoute("UserDetail")
        .attachPatternMatched(this._onObjectMatched, this);
    }
  }

  /**
   * Get date from Global filter and format it
   * because date from DateRange filter is object
   * START
   **/
  private getGlobalDateFilter(): Filter[] {
    const { from, to } = this.getGlobalDateRange();

    return [
      new Filter({
        path: "LoginDate",
        operator: FilterOperator.BT,
        value1: from,
        value2: to,
      }),
    ];
  }

  private getGlobalDateRange() {
    const oComponent = this.getAppComponent();
    const oGlobalModel = oComponent?.getModel("global") as JSONModel;

    if (!oGlobalModel) {
      return { from: "", to: "" };
    }

    const oFrom = oGlobalModel.getProperty("/fromDate");
    const oTo = oGlobalModel.getProperty("/toDate");

    const formatDate = (d: Date) => d.toISOString().split("T")[0];

    return {
      from: formatDate(oFrom),
      to: formatDate(oTo),
    };
  }
  /**
   * END
   **/

  /**
   * Handles route matching for AuthDetail page
   * and loads detail data based on the navigation key.
   **/
  private async _onObjectMatched(oEvent: any): Promise<void> {
    // Get parameter from URL
    const sUsername = oEvent.getParameter("arguments").username;

    const oView = this.getView();
    if (!oView || !sUsername) return;

    oView.setBusy(true);

    try {
      await Promise.all([
        this._loadUserDetail(sUsername),
        this._loadUserLogs(sUsername),
        this._loadUserAuthLogPerDay(sUsername),
        this._loadUserActivity(sUsername),
      ]);
    } catch (oError) {
      MessageBox.error("Failed to load user data. Please try again.");
    } finally {
      oView.setBusy(false);
    }
  }

  /**
   * Called when the View has been rendered (so its HTML is part of the document). Post-rendering manipulations of the HTML could be done here.
   * This hook is the same one that SAPUI5 controls get after being rendered.
   * @memberOf useraudit.controller.UserDetail
   */
  public onAfterRendering(): void {
    const { from, to } = this.getGlobalDateRange();

    // Set range in date picker
    const oDatePicker = this.byId("AuthDatePickerId") as DatePicker;

    if (!oDatePicker) return;

    oDatePicker.setMinDate(new Date(from));
    oDatePicker.setMaxDate(new Date(to));
  }

  /**
   * Load user detail information
   **/
  private async _loadUserDetail(sUsername: string): Promise<void> {
    const oView = this.getView();
    const oModel = (this as any).getAppComponent().getModel() as ODataModel;

    // Create a list binding to /UserDetail with $filter
    const oUserDetai = oModel.bindList("/UserDetail", undefined, undefined, [
      new Filter("UserName", FilterOperator.EQ, sUsername),
    ]) as ODataListBinding;

    // Executes the OData call
    const aContexts = await oUserDetai.requestContexts(0, 1);

    // Set data into model
    if (aContexts.length > 0) {
      const oData = aContexts[0].getObject();
      const oUserDetailModel = new JSONModel(oData);

      oView?.setModel(oUserDetailModel, "UserDetailData");
    } else {
      MessageBox.error("Failed to load user detail data. Please try again.");
    }
  }

  /**
   * Load user auth log
   **/
  private async _loadUserLogs(sUsername: string): Promise<void> {
    const oUserAuthLogPerDayData = {} as any;

    const oModel = (this as any).getAppComponent().getModel() as ODataModel;

    // Create a list binding to /AuthLogChartByUser with $filter
    const oUserAuthChart = oModel.bindList(
      "/AuthLogChartByUser",
      undefined,
      undefined,
      [
        new Filter("Username", FilterOperator.EQ, sUsername),
        ...this.getGlobalDateFilter(),
      ],
    ) as ODataListBinding;

    // Executes the OData call
    const aContextsChart = await oUserAuthChart.requestContexts();

    const aDataChart = aContextsChart.map((oContext) => oContext.getObject());

    // Group by + SUM data
    aDataChart.forEach((oContext) => {
      const key = oContext.LoginResult;

      if (!oUserAuthLogPerDayData[key]) {
        oUserAuthLogPerDayData[key] = {
          LoginResult: oContext.LoginResult,
          CountLoginLog: 0,
        };
      }

      oUserAuthLogPerDayData[key].CountLoginLog += oContext.CountLoginLog;
    });

    //  Convert into array
    let aUserAuthLogPerDayData = Object.values(oUserAuthLogPerDayData);

    const oJsonModel = new JSONModel(aUserAuthLogPerDayData);

    // Set data into Model AuthLogChartByUser
    this.getView()?.setModel(oJsonModel, "AuthLogChartByUserData");

    // Create a list binding to /UserAuthLog with $filter
    const oUserAuthTable = oModel.bindList(
      "/UserAuthLog",
      undefined,
      undefined,
      [
        new Filter("Username", FilterOperator.EQ, sUsername),
        ...this.getGlobalDateFilter(),
      ],
    ) as ODataListBinding;

    const aContextsTable = await oUserAuthTable.requestContexts();
    const aDataTable = aContextsTable.map((oContext) => oContext.getObject());

    const oTableModel = new JSONModel(aDataTable);
    this.getView()?.setModel(oTableModel, "UserAuthLogData");
  }

  /**
   * Load User Auth Log Per Day
   **/
  private async _loadUserAuthLogPerDay(sUsername: string): Promise<void> {
    const oLogUserPerDay = {} as any;

    const oModel = (this as any).getAppComponent().getModel() as ODataModel;

    // Create a list binding to /UserAuthLogPerDay with $filter
    const oUserAuthLogPerDay = oModel.bindList(
      "/UserAuthLogPerDay",
      undefined,
      undefined,
      [
        new Filter("UserName", FilterOperator.EQ, sUsername),
        ...this.getGlobalDateFilter(),
      ],
    ) as ODataListBinding;

    const aContextsLogPerDay = await oUserAuthLogPerDay.requestContexts();

    aContextsLogPerDay.forEach((oContext) => {
      const oObj = oContext.getObject();

      const key = oObj.LoginDate;

      if (!oLogUserPerDay[key]) {
        oLogUserPerDay[key] = {
          LoginDate: oObj.LoginDate,
          TotalLoginCount: 0,
        };
      }

      oLogUserPerDay[key].TotalLoginCount += oObj.TotalLoginCount;
    });

    //  Convert into array
    let aLogUserPerDay = Object.values(oLogUserPerDay);

    const oLogPerDaModel = new JSONModel(aLogUserPerDay);

    this.getView()?.setModel(oLogPerDaModel, "UserAuthLogPerDayData");
  }

  /**
   * Load user activity information
   **/
  private async _loadUserActivity(sUsername: string): Promise<void> {
    const { from, to } = this.getGlobalDateRange();

    const oTCodePerUserData = {} as any;

    const oModel = (this as any).getAppComponent().getModel() as ODataModel;

    // Create a list binding to /ActivityTCodeByUser with $filter
    const oActivityTCodeByUser = oModel.bindList(
      "/ActivityTCodeByUser",
      undefined,
      undefined,
      [
        new Filter("Username", FilterOperator.EQ, sUsername),
        new Filter({
          path: "ActivityDate",
          operator: FilterOperator.BT,
          value1: from,
          value2: to,
        }),
      ],
    ) as ODataListBinding;

    // Executes the OData call
    const aContextsActivityTCodeByUser =
      await oActivityTCodeByUser.requestContexts();

    // Group by + SUM data
    aContextsActivityTCodeByUser.forEach((oContext) => {
      const oObj = oContext.getObject();

      const key = oObj.TCode;

      // If TCode is exist, plus TCodeCount, else create new obj
      if (!oTCodePerUserData[key]) {
        oTCodePerUserData[key] = {
          TCode: oObj.TCode,
          TCodeName: oObj.TCodeName,
          TCodeCount: 0,
        };
      }

      oTCodePerUserData[key].TCodeCount += oObj.TCodeCount;
    });

    //  Convert into array
    let aTCodePerUserData = Object.values(oTCodePerUserData);

    // Sort data
    aTCodePerUserData.sort((a: any, b: any) => b.TCodeCount - a.TCodeCount);

    //  Top 5
    aTCodePerUserData = aTCodePerUserData.slice(0, 5);

    //  Add label
    aTCodePerUserData.forEach((item: any) => {
      item.Label = `${item.TCode} - ${item.TCodeName}`;
    });

    // Set data into Model ActivityTCodeByUser
    const oJsonModelActTcode = new JSONModel(aTCodePerUserData);
    this.getView()?.setModel(oJsonModelActTcode, "TCodeByUserData");

    // Create a list binding to /UserActivityLog with $filter
    const oUserActTable = oModel.bindList(
      "/UserActivityLog",
      undefined,
      undefined,
      [
        new Filter("Username", FilterOperator.EQ, sUsername),
        new Filter({
          path: "ActivityDate",
          operator: FilterOperator.BT,
          value1: from,
          value2: to,
        }),
      ],
    ) as ODataListBinding;

    const aContextsTable = await oUserActTable.requestContexts();
    const aDataTable = aContextsTable.map((oContext) => oContext.getObject());

    const oTableModel = new JSONModel(aDataTable);
    this.getView()?.setModel(oTableModel, "UserActivityLogData");
  }

  /**
   * Navigate to AuthDetail page
   **/
  public onNavigationToAuthDetail(oEvent: any): void {
    // Get control and BindingContext of line
    const oItem = oEvent.getSource();
    const oContext = oItem.getBindingContext("UserAuthLogData");

    if (oContext) {
      const sSessionId = oContext.getProperty("SessionId");

      // Navigate with parameter session_id
      const oRouter = (this as any).getAppComponent().getRouter();
      if (oRouter) {
        oRouter.navTo("AuthDetail", {
          key: sSessionId,
        });
      }
    }
  }

  /**
   * Called when the user use filter auth status
   **/
  public onFilterAuth(): void {
    this.applyFilters();
  }

  /**
   * Called when the user use filter auth date
   **/
  public onFilterAuthLoginDate(): void {
    this.applyFilters();
  }

  /**
   * Execute logic search and filter
   **/
  public applyFilters(): void {
    const aFilters: Filter[] = [];

    const oTable = this.byId("AuthTableId") as Table;
    const oBinding = oTable.getBinding("rows") as ODataListBinding;

    // Get value from search and select
    const sStatus = (
      this.byId("AuthStatusSelectId") as Select
    ).getSelectedKey();

    if (sStatus) {
      aFilters.push(new Filter("LoginResult", FilterOperator.EQ, sStatus));
    }

    // Get value select date picker
    const oDatePicker = this.byId("AuthDatePickerId") as DatePicker;
    if (oDatePicker) {
      const oDate = oDatePicker.getDateValue();

      if (oDate) {
        const oFormatter = DateFormat.getDateInstance({
          pattern: "yyyy-MM-dd",
        });

        const sDate = oFormatter.format(oDate);

        aFilters.push(new Filter("LoginDate", FilterOperator.EQ, sDate));
      }
    }

    // Final filter and render into table
    const aFinalFilters = oDatePicker
      ? aFilters
      : [...this.getGlobalDateFilter(), ...aFilters];

    if (oBinding) {
      oBinding.filter(aFinalFilters);
    }
  }
  //  Exports Excel file
  public onExportUserDetailExcel(): void {
    const sFileName = `User_${this._sFromDate}_to_${this._sToDate}.xlsx`;

    MessageBox.confirm("Do you want to export this data to Excel?", {
      title: "Confirm Export",
      actions: ["YES", "NO"],
      emphasizedAction: "YES",

      onClose: (oAction: string | null) => {
        if (oAction === "YES") {
          const oModel = this.getView()?.getModel(
            "UserAuthLogData",
          ) as JSONModel;
          const aData = oModel?.getData();

          if (!aData || aData.length === 0) {
            MessageBox.error("No data to export.");
            return;
          }

          const aCols = [
            { label: "User Session", property: "SessionId", width: 25 },
            { label: "User Name", property: "Username", width: 15 },
            { label: "Login Result", property: "LoginResult", width: 10 },
            { label: "Login Date", property: "LoginDate", width: 15 },
            { label: "Login Time", property: "LoginTime", width: 15 },
            { label: "Login Message", property: "LoginMessage", width: 150 },
            { label: "Logout Date", property: "LogoutDate", width: 15 },
            { label: "Logout Time", property: "LogoutTime", width: 15 },
            { label: "Event ID", property: "EventId", width: 10 },
          ];

          const oSettings = {
            workbook: { columns: aCols },
            dataSource: aData,
            fileName: sFileName,
            worker: false,
          };

          const oSheet = new Spreadsheet(oSettings);

          oSheet
            .build()
            .then(() => {
              MessageToast.show("Export successful!", { duration: 3000 });
            })
            .catch(() => {
              MessageBox.error("Export failed.");
            })
            .finally(() => {
              oSheet.destroy();
            });
        }
      },
    });
  }
}

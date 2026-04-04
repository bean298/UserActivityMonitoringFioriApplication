import MessageBox from "sap/m/MessageBox";
import Controller from "sap/fe/core/PageController";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import JSONModel from "sap/ui/model/json/JSONModel";
import ODataListBinding from "sap/ui/model/odata/v4/ODataListBinding";
import ODataModel from "sap/ui/model/odata/v4/ODataModel";
import Formatter from "useraudit/formatter/Formatter";
import Spreadsheet from "sap/ui/export/Spreadsheet";
import MessageToast from "sap/m/MessageToast";
import DateFormat from "sap/ca/ui/model/format/DateFormat";

export default class UserDetail extends Controller {
  public formatter = Formatter;

  private _sFromDate!: string;
  private _sToDate!: string;
  private _aDefaultFilters: Filter[] = [];
  private _aUserFilters: Filter[] = [];
  private _aUserDateFilters: Filter[] = [];
  private _sCurrentUsername: string = "";

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
   * Handles route matching for AuthDetail page
   * and loads detail data based on the navigation key.
   **/
  private async _onObjectMatched(oEvent: any): Promise<void> {
    // Get parameter from URL
    const sUsername = oEvent.getParameter("arguments").username;
    this._sCurrentUsername = sUsername;

    const oView = this.getView();
    if (!oView || !sUsername) return;

    // Take a last 6 days
    const oToday = new Date();
    const o7DaysAgo = new Date();
    o7DaysAgo.setDate(oToday.getDate() - 3);

    const formatDate = (oDate: Date) => {
      return oDate.toISOString().split("T")[0];
    };

    this._sFromDate = formatDate(o7DaysAgo);
    this._sToDate = formatDate(oToday);

    this._aDefaultFilters = [
      new Filter({
        path: "LoginDate",
        operator: FilterOperator.BT,
        value1: this._sFromDate,
        value2: this._sToDate,
      }),
    ];
    // Reset filter
    this._aUserFilters = [];
    this._aUserDateFilters = [];
    oView.setBusy(true);

    try {
      await Promise.all([
        this._loadUserDetail(sUsername),
        this._loadUserLogs(sUsername),
        this._loadUserActivity(sUsername),
        this._loadUserAuthLogPerDay(sUsername),
      ]);
    } catch (oError) {
      MessageBox.error("Failed to load user data. Please try again.");
    } finally {
      oView.setBusy(false);
    }
  }
  public async onFilterAuth(): Promise<void> {
    const oView = this.getView();
    if (!oView) return;

    oView.setBusy(true);

    // Get Status from Select
    const sStatus = (this.byId("AuthStatusSelectId") as any).getSelectedKey();
    this._aUserFilters =
      sStatus && sStatus !== ""
        ? [new Filter("LoginResult", FilterOperator.EQ, sStatus)]
        : [];

    // Get Date from DatePicker
    const oDatePicker = this.byId("AuthDatePickerId") as any;
    const oDate = oDatePicker.getDateValue();

    if (oDate) {
      const sDate = DateFormat.getDateInstance({
        pattern: "yyyy-MM-dd",
      }).format(oDate, false);
      this._aUserDateFilters = [
        new Filter("LoginDate", FilterOperator.EQ, sDate),
      ];
    } else {
      this._aUserDateFilters = [];
    }

    // Call the function to reload the data
    try {
      await this._loadUserLogs(this._sCurrentUsername);
    } finally {
      oView.setBusy(false);
    }
  }
  /**
   * Load user detail information
   **/
  private async _loadUserDetail(sUsername: string): Promise<void> {
    const oView = this.getView();
    const oModel = (this as any).getAppComponent().getModel() as ODataModel;
    const aDateFilters =
      this._aUserDateFilters.length > 0
        ? this._aUserDateFilters
        : this._aDefaultFilters;
    const aFinalFilters = [
      new Filter("Username", FilterOperator.EQ, sUsername),
      ...aDateFilters,
      ...this._aUserFilters,
    ];
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

    // --- SETUP FILTERS ---
    const aDateFilters =
      this._aUserDateFilters.length > 0
        ? this._aUserDateFilters
        : this._aDefaultFilters;
    const aFinalFilters = [
      new Filter("Username", FilterOperator.EQ, sUsername),
      ...aDateFilters,
      ...this._aUserFilters,
    ];
    // Create a list binding to /AuthLogChartByUser with $filter
    const oUserAuthChart = oModel.bindList(
      "/AuthLogChartByUser",
      undefined,
      undefined,
      aFinalFilters,
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
      aFinalFilters,
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
    const aDateFilters =
      this._aUserDateFilters.length > 0
        ? this._aUserDateFilters
        : this._aDefaultFilters;
    const aFinalFilters = [
      new Filter("UserName", FilterOperator.EQ, sUsername),
      ...aDateFilters,
      ...this._aUserFilters,
    ];
    // Create a list binding to /UserAuthLogPerDay with $filter
    const oUserAuthLogPerDay = oModel.bindList(
      "/UserAuthLogPerDay",
      undefined,
      undefined,
      aFinalFilters,
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
          value1: this._sFromDate,
          value2: this._sToDate,
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
          value1: this._sFromDate,
          value2: this._sToDate,
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

  //  Exports Excel file.
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

import MessageBox from "sap/m/MessageBox";
import Controller from "sap/fe/core/PageController";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import JSONModel from "sap/ui/model/json/JSONModel";
import ODataListBinding from "sap/ui/model/odata/v4/ODataListBinding";
import ODataModel from "sap/ui/model/odata/v4/ODataModel";
import Formatter from "useraudit/formatter/Formatter";

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
        this._loadUserActivity(sUsername),
      ]);
    } catch (oError) {
      MessageBox.error("Failed to load user data. Please try again.");
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
    const oModel = (this as any).getAppComponent().getModel() as ODataModel;

    // Create a list binding to /AuthLogChartByUser with $filter
    const oUserAuthChart = oModel.bindList(
      "/AuthLogChartByUser",
      undefined,
      undefined,
      [new Filter("Username", FilterOperator.EQ, sUsername)],
    ) as ODataListBinding;

    // Executes the OData call
    const aContextsChart = await oUserAuthChart.requestContexts();

    const aDataChart = aContextsChart.map((oContext) => oContext.getObject());

    const oJsonModel = new JSONModel(aDataChart);

    // Set data into Model AuthLogChartByUser
    this.getView()?.setModel(oJsonModel, "AuthLogChartByUserData");

    // Create a list binding to /UserAuthLog with $filter
    const oUserAuthTable = oModel.bindList(
      "/UserAuthLog",
      undefined,
      undefined,
      [new Filter("Username", FilterOperator.EQ, sUsername)],
    ) as ODataListBinding;

    const aContextsTable = await oUserAuthTable.requestContexts();
    const aDataTable = aContextsTable.map((oContext) => oContext.getObject());

    const oTableModel = new JSONModel(aDataTable);
    this.getView()?.setModel(oTableModel, "UserAuthLogData");
  }

  /**
   * Load user activity information
   **/
  private async _loadUserActivity(sUsername: string): Promise<void> {
    const oModel = (this as any).getAppComponent().getModel() as ODataModel;

    // // Create a list binding to /UserActivityLog with $filter
    // const oUserAuthChart = oModel.bindList(
    //   "/UserActivityLog",
    //   undefined,
    //   undefined,
    //   [new Filter("Username", FilterOperator.EQ, sUsername)],
    // ) as ODataListBinding;

    // // Executes the OData call
    // const aContextsChart = await oUserAuthChart.requestContexts();

    // const aDataChart = aContextsChart.map((oContext) => oContext.getObject());

    // const oJsonModel = new JSONModel(aDataChart);

    // // Set data into Model AuthLogChartByUser
    // this.getView()?.setModel(oJsonModel, "AuthLogChartByUserData");

    // Create a list binding to /UserActivityLog with $filter
    const oUserActTable = oModel.bindList(
      "/UserActivityLog",
      undefined,
      undefined,
      [new Filter("Username", FilterOperator.EQ, sUsername)],
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
}

import Controller from "sap/fe/core/PageController";
import ODataListBinding from "sap/ui/model/odata/v4/ODataListBinding";
import ODataModel from "sap/ui/model/odata/v4/ODataModel";
import JSONModel from "sap/ui/model/json/JSONModel";
import Input from "sap/m/Input";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import Select from "sap/m/Select";
import Dialog from "sap/m/Dialog";
import Formatter from "useraudit/formatter/Formatter";
import MessageBox from "sap/m/MessageBox";
import Spreadsheet from "sap/ui/export/Spreadsheet";
import MessageToast from "sap/m/MessageToast";
import Fragment from "sap/ui/core/Fragment";
import DateFormat from "sap/ui/core/format/DateFormat";
import DatePicker from "sap/m/DatePicker";

export default class Main extends Controller {
  public formatter = Formatter;

  // ===== View Settings Feature =====
  private _oViewSettingsDialog: Dialog | null = null;
  private _oUserSearchHelpDialog: Dialog | null = null;

  /**
   * Called when the controller is initialized.
   **/
  public onInit(): void {
    super.onInit();

    // Create view model
    const oOverviewModel = new JSONModel({
      totalUsers: 0,
      totalLogs: 0,
      successLogin: 0,
      failedLogin: 0,
      dumpCount: 0,
    });
    this.getView()?.setModel(oOverviewModel, "Overview");

    this.onInitCount();
    this.onInitLogCount();
    this.onInitTcodeCount();
    this.onInitDumpCount();
    this.onInitOverviewData();
  }

  /**
   * Fetches the total number of records from the UserAuthLog entity
   **/
  public async onInitCount(): Promise<void> {
    try {
      // Create view model
      const oViewModel = new JSONModel({
        count: 0,
      });
      this.getView()?.setModel(oViewModel, "view");

      // Get OData V4 model from the App Component
      const oModel = this.getAppComponent().getModel() as ODataModel;

      // Get model Overview
      const oOverviewModel = this.getView()?.getModel("Overview") as JSONModel;

      // Create a list binding to /UserAuthLog with $count enabled
      const oBinding = oModel.bindList(
        "/UserAuthLog",
        undefined,
        undefined,
        undefined,
        { $count: true },
      ) as ODataListBinding;

      // Executes the OData call
      await oBinding.requestContexts();

      const iCount = oBinding.getLength();

      // Create property of view model
      oViewModel.setProperty("/count", iCount);

      oOverviewModel.setProperty("/totalLogs", iCount);
    } catch (error) {
      MessageBox.error("Failed to load chart data.");
    }
  }

  /**
   * Fetches the data for Overview Data
   **/
  public async onInitOverviewData(): Promise<void> {
    try {
      // Get OData V4 model from the App Component
      const oModel = this.getAppComponent().getModel() as ODataModel;

      // Get model Overview
      const oOverviewModel = this.getView()?.getModel("Overview") as JSONModel;

      // Create a list binding to /UserSearchHelp
      const oBindingTotalUser = oModel.bindList(
        "/UserSearchHelp",
        undefined,
        undefined,
        undefined,
        {
          $count: true,
        },
      ) as ODataListBinding;

      // Executes the OData call
      await oBindingTotalUser.requestContexts(0, 1);

      const iTotalUsers = oBindingTotalUser.getLength();

      // Create property of view model
      oOverviewModel.setProperty("/totalUsers", iTotalUsers);

      // Create a list binding to /UserActivityLog
      const oBindingTotalDump = oModel.bindList(
        "/UserActivityLog",
        undefined,
        undefined,
        [new Filter("ActivityType", FilterOperator.EQ, "DUMP")],
        {
          $count: true,
        },
      ) as ODataListBinding;

      // Executes the OData call
      await oBindingTotalDump.requestContexts();

      const iDumpCount = oBindingTotalDump.getLength();

      // Create property of view model
      oOverviewModel.setProperty("/dumpCount", iDumpCount);
    } catch (error) {
      MessageBox.error("Failed to load overview data.");
    }
  }

  /**
   * Fetches the data of UserAuthLogChart entity
   **/
  public async onInitLogCount(): Promise<void> {
    try {
      // Get OData V4 model from the App Component
      const oModel = this.getAppComponent().getModel() as ODataModel;

      // Get model Overview
      const oOverviewModel = this.getView()?.getModel("Overview") as JSONModel;

      // Create a list binding to /UserAuthLogChart
      const oBinding = oModel.bindList("/UserAuthLogChart") as ODataListBinding;

      // Executes the OData call
      const aContexts = await oBinding.requestContexts();

      const aData = aContexts.map((oContext) => oContext.getObject());

      const oJsonModel = new JSONModel(aData);

      // Set data for overview
      aData.forEach((item) => {
        if (item.LoginResult === "SUCCESS") {
          oOverviewModel.setProperty("/successLogin", item.CountLoginLog);
        } else if (item.LoginResult === "FAIL") {
          oOverviewModel.setProperty("/failedLogin", item.CountLoginLog);
        }
      });

      // Set data into Model authLogChart
      this.getView()?.setModel(oJsonModel, "authLogChart");
    } catch (error) {
      MessageBox.error("Failed to load Login data.");
    }
  }

  /**
   * Fetches the data of ActivityLogChart entity
   **/
  public async onInitTcodeCount(): Promise<void> {
    try {
      // Get OData V4 model from the App Component
      const oModel = this.getAppComponent().getModel() as ODataModel;

      // Create a list binding to /ActivityLogChart
      const oBinding = oModel.bindList(
        "/ActivityLogChart",
        undefined,
        undefined,
        undefined,
        {
          $orderby: "TCodeCount desc",
        },
      ) as ODataListBinding;

      // Executes the OData call
      const aContexts = await oBinding.requestContexts(0, 5);

      const aData = aContexts.map((oContext) => {
        // Create label field
        const oObj = oContext.getObject();
        oObj.Label = `${oObj.TCode} - ${oObj.TCodeName}`;

        return oObj;
      });

      const oJsonModel = new JSONModel(aData);

      // Set data into Model tCodeChart
      this.getView()?.setModel(oJsonModel, "tCodeChart");
    } catch (error) {
      MessageBox.error("Failed to load TCode data.");
    }
  }

  /**
   * Fetches the data of DumpActivityChart entity
   **/
  public async onInitDumpCount(): Promise<void> {
    try {
      // Get OData V4 model from the App Component
      const oModel = this.getAppComponent().getModel() as ODataModel;

      // Create a list binding to /DumpActivityChart
      const oBinding = oModel.bindList(
        "/DumpActivityChart",
        undefined,
        undefined,
        undefined,
        {
          $orderby: "DumpCount desc",
        },
      ) as ODataListBinding;

      // Executes the OData call
      const aContexts = await oBinding.requestContexts(0, 10);

      const aData = aContexts.map((oContext) => oContext.getObject());

      const oJsonModel = new JSONModel(aData);

      // Set data into Model tCodeChart
      this.getView()?.setModel(oJsonModel, "dumpChart");
    } catch (error) {
      MessageBox.error("Failed to load Dump data.");
    }
  }

  /**
   * Called when the user use search
   **/
  public onSearchUserName(): void {
    this.applyFilters();
  }

  /**
   * Called when the user use filter status
   **/
  public onFilterStatus(): void {
    this.applyFilters();
  }

  /**
   * Called when the user use filter date
   **/
  public onFilterLoginDate(): void {
    this.applyFilters();
  }

  /**
   * Execute logic search and filter
   **/
  public applyFilters(): void {
    const aFilters: Filter[] = [];

    // Get table and its OData list binding
    const oTable = this.byId("maiTableId");
    const oBinding = oTable?.getBinding("rows") as ODataListBinding;

    // Get value from search and select
    const sSearch = (this.byId("userSearchId") as Input).getValue();
    if (sSearch) {
      aFilters.push(new Filter("Username", FilterOperator.Contains, sSearch));
    }

    const sStatus = (
      this.byId("mainHeaderSelectId") as Select
    ).getSelectedKey();

    if (sStatus) {
      aFilters.push(new Filter("LoginResult", FilterOperator.EQ, sStatus));
    }

    // Get value select date picker
    const oDatePicker = this.byId("mainDatePickerId") as DatePicker;

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

    oBinding.filter(aFilters);
  }

  /**
   * Exports the currently bound table data to an Excel file.
   * Uses sap.ui.export.Spreadsheet to generate the file
   * based on the current OData V4 list binding.
   */
  public onExportExcel(): void {
    MessageBox.confirm("Do you want to export this data to Excel?", {
      title: "Confirm Export",
      actions: ["YES", "NO"],
      emphasizedAction: "YES",

      onClose: (oAction: string | null) => {
        if (oAction === "YES") {
          // Get table and its OData list binding
          const oTable = this.byId("maiTableId");
          const oBinding = oTable?.getBinding("rows") as ODataListBinding;
          // Format in Excel
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

          // Spreadsheet config
          const oSettings = {
            workbook: { columns: aCols },
            dataSource: oBinding,
            fileName: "UserAuthenticationLogs.xlsx",
            worker: false,
          };

          // Spreadsheet config
          const oSheet = new Spreadsheet(oSettings);
          // Create Excel file and download it
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

  /**
   * Called when the value help of the user search input is triggered.
   * Fetches data from the OData service
   **/
  public async onUserSearchHelp(): Promise<void> {
    try {
      // Get OData V4 model from the App Component
      const oModel = this.getAppComponent().getModel() as ODataModel;

      // Create a list binding to /UserSearchHelp
      const oBinding = oModel.bindList("/UserSearchHelp") as ODataListBinding;

      // Executes the OData call and load data
      const aContexts = await oBinding.requestContexts();
      const aData = aContexts.map((oContent) => oContent.getObject());

      const oJsonModel = new JSONModel(aData);

      this.getView()?.setModel(oJsonModel, "userSeachHelp");

      if (!this._oUserSearchHelpDialog) {
        // Load fragment
        this._oUserSearchHelpDialog = (await Fragment.load({
          name: "useraudit.fragment.UserSearchHelp",
          controller: this,
        })) as Dialog;

        this.getView()?.addDependent(this._oUserSearchHelpDialog);
      }

      this._oUserSearchHelpDialog.open();
    } catch (error) {
      MessageBox.error("Failed to load user search help.");
    }
  }

  /**
   * Select User SearchHelp
   **/
  public onUserSelect(oEvent: any): void {
    const oItem = oEvent.getParameter("listItem");

    const oContext = oItem.getBindingContext("userSeachHelp");
    const oSelected = oContext?.getObject();

    if (oSelected) {
      (this.byId("userSearchId") as Input).setValue(oSelected.Username);
    }

    this.applyFilters();

    this._oUserSearchHelpDialog?.close();
  }

  /**
   * Close Fragment User Search Help
   **/
  public onCloseUserDialog(): void {
    this._oUserSearchHelpDialog?.close();
  }

  /**
   * Navigate to AuthDetail page
   **/
  public onPressNavigate(oEvent: any): void {
    // Get control and BindingContext of line
    const oItem = oEvent.getSource();
    const oContext = oItem.getBindingContext();

    if (oContext) {
      // Get path
      let sPath = oContext.getPath();

      if (sPath.startsWith("/")) {
        sPath = sPath.substring(1);
      }

      // Navigate with parameter session_id
      const oRouter = (this as any).getAppComponent().getRouter();
      if (oRouter) {
        oRouter.navTo("AuthDetail", {
          key: sPath,
        });
      }
    }
  }

  /**
   * Open Fragment Settings Columns
   **/
  public async onOpenViewSettings(): Promise<void> {
    if (!this._oViewSettingsDialog) {
      // Load fragment
      this._oViewSettingsDialog = (await Fragment.load({
        name: "useraudit.fragment.ViewSettingsDialog",
        controller: this,
      })) as Dialog;

      // Add Fragment into view
      this.getView()?.addDependent(this._oViewSettingsDialog);

      this.initializeColumnModel();
    }

    // Open
    this._oViewSettingsDialog.open();
  }

  /**
   * Read columns list of table -> JSONModel
   * Bring it into Dialog View
   **/
  public initializeColumnModel(): void {
    // Get table and its colums
    const oTable = this.byId("maiTableId") as any;
    const aColumns = oTable.getColumns();

    // Remove Navigator column
    const aFilteredColumns = aColumns.filter(
      (oColumn: any) => !oColumn.getId().includes("columnNavigator"),
    );

    // Create new array contain every column object
    const aColumnData = aFilteredColumns.map((oColumn: any) => {
      const oLabel = oColumn.getLabel();

      return {
        id: oColumn.getId(),
        label: oLabel ? oLabel.getText() : oColumn.getId(),
        visible: oColumn.getVisible(),
      };
    });

    // Create JSONModel, property columns
    const oModel = new JSONModel({
      columns: aColumnData,
    });

    this.getView()?.setModel(oModel, "columnsModel");
  }

  /**
   * Open Fragment Settings Columns
   **/
  public onConfirmViewSettings(): void {
    // Get table and its colums
    const oTable = this.byId("maiTableId") as any;
    const aColumns = oTable.getColumns();

    // Get model and property
    const oModel = this.getView()?.getModel("columnsModel") as JSONModel;
    const aData = oModel.getProperty("/columns");

    // Set visible for column
    aColumns.forEach((oColumn: any) => {
      const oMatch = aData.find((column: any) => column.id === oColumn.getId());

      if (oMatch) {
        oColumn.setVisible(oMatch.visible);
      }
    });

    this._oViewSettingsDialog?.close();
  }

  /**
   * Close Fragment Settings Columns
   **/
  public onCancelViewSettings(): void {
    this._oViewSettingsDialog?.close();
  }

  /**
   * Refresh data of table
   **/
  public async onRefreshTable(): Promise<void> {
    const oTable = this.byId("maiTableId") as any;
    const oBinding = oTable.getBinding("rows") as ODataListBinding;

    oTable.setBusy(true);

    try {
      // Refresh data
      await oBinding?.requestRefresh();

      // Refresh dashboard data
      await this.onInitCount();
      await this.onInitLogCount();
      await this.onInitTcodeCount();

      MessageToast.show("Data refreshed");
    } finally {
      oTable.setBusy(false);
    }
  }

  /**
   * Navigate to user detail page
   **/
  public onNavigateUserDetail(oEvent: any): void {
    // Get control and BindingContext of line
    const oItem = oEvent.getSource();
    const oContext = oItem.getBindingContext();

    if (oContext) {
      const sUsername = oContext.getProperty("Username");

      // Navigate with parameter username
      const oRouter = (this as any).getAppComponent().getRouter();
      if (oRouter) {
        oRouter.navTo("UserDetail", {
          username: sUsername,
        });
      }
    }
  }
}

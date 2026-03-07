import Controller from "sap/fe/core/PageController";
import ODataListBinding from "sap/ui/model/odata/v4/ODataListBinding";
import ODataModel from "sap/ui/model/odata/v4/ODataModel";
import JSONModel from "sap/ui/model/json/JSONModel";
import Input from "sap/m/Input";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import Select from "sap/m/Select";
import List from "sap/m/List";
import StandardListItem from "sap/m/StandardListItem";
import Dialog from "sap/m/Dialog";
import Button from "sap/m/Button";
import Formatter from "useraudit/formatter/Formatter";
import MessageBox from "sap/m/MessageBox";
import DateRangeSelection from "sap/m/DateRangeSelection";
import Spreadsheet from "sap/ui/export/Spreadsheet";
import MessageToast from "sap/m/MessageToast";
import Fragment from "sap/ui/core/Fragment";

export default class Main extends Controller {
  public formatter = Formatter;

  // ===== View Settings Feature =====
  private _oViewSettingsDialog: Dialog | null = null;

  /**
   * Called when the controller is initialized.
   **/
  public onInit(): void {
    super.onInit();

    this.onInitCount();

    this.onInitLogCount();
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
    } catch (error) {
      MessageBox.error("Failed to load chart data.");
    }
  }

  /**
   * Fetches the data of UserAuthLogChart entity
   **/
  public async onInitLogCount(): Promise<void> {
    try {
      // Get OData V4 model from the App Component
      const oModel = this.getAppComponent().getModel() as ODataModel;

      // Create a list binding to /UserAuthLogChart
      const oBinding = oModel.bindList("/UserAuthLogChart") as ODataListBinding;

      // Executes the OData call
      const aContexts = await oBinding.requestContexts();

      const aData = aContexts.map((oContext) => oContext.getObject());

      const oJsonModel = new JSONModel(aData);

      // Set data into Model authLogChart
      this.getView()?.setModel(oJsonModel, "authLogChart");
    } catch (error) {
      MessageBox.error("Failed to load chart data.");
    }
  }

  /**
   * Called when the user use search
   **/
  public onSearchUserName(): void {
    this.applyFilters();
  }

  /**
   * Called when the user use filter
   **/
  public onFilterStatus(): void {
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

    // Get value select date range
    const oDateRange = this.byId("mainDateRangeId") as DateRangeSelection;

    if (oDateRange) {
      const oFromDate = oDateRange.getDateValue();
      const oToDate = oDateRange.getSecondDateValue();

      if (oFromDate && oToDate) {
        aFilters.push(
          new Filter("LoginDate", FilterOperator.BT, oFromDate, oToDate),
        );
      }
    }

    oBinding.filter(aFilters);
  }

  /**
   * Triggered when the user changes the "Rows per page".
   */
  public onRowCountChange(oEvent: any): void {
    const sKey = oEvent.getParameter("selectedItem").getKey();
    const oTable = this.byId("maiTableId") as any;
    if (oTable) {
      oTable.setVisibleRowCount(parseInt(sKey));
    }
  }

  /**
   * Exports the currently bound table data to an Excel file.
   * Uses sap.ui.export.Spreadsheet to generate the file
   * based on the current OData V4 list binding.
   */
  public onExportExcel(): void {
    // Get table and its OData list binding
    const oTable = this.byId("maiTableId");
    const oBinding = oTable?.getBinding("rows") as ODataListBinding;

    // Format in Excel
    const aCols = [
      {
        label: "User Session",
        property: "SessionId",
        width: 25,
      },
      { label: "User Name", property: "Username", width: 15 },
      {
        label: "Login Result",
        property: "LoginResult",
        width: 10,
      },
      { label: "Login Date", property: "LoginDate", width: 15 },
      { label: "Login Time", property: "LoginTime", width: 15 },
      {
        label: "Login Message",
        property: "LoginMessage",
        width: 150,
      },
      { label: "Logout Date", property: "LogoutDate", width: 15 },
      { label: "Logout Time", property: "LogoutTime", width: 15 },
      { label: "Event ID", property: "EventId", width: 10 },
    ];

    // Spreadsheet config
    const oSettings = {
      workbook: {
        columns: aCols,
      },
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
        MessageToast.show("Export successful!", {
          duration: 3000,
        });
      })
      .catch(() => {
        MessageBox.error("Export failed.");
      })
      .finally(() => {
        oSheet.destroy();
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

      // Create List
      const oList = new List({
        mode: "SingleSelectMaster",
        items: {
          path: "/",
          template: new StandardListItem({
            title: "{Username}",
          }),
        },
        selectionChange: (oEvent) => {
          //Get line that user click
          const oItem = oEvent.getParameter("listItem");
          if (!oItem) {
            return;
          }

          // Connects the UI item to its underlying model data
          const oContext = oItem.getBindingContext();

          // Get the actual data object from the model
          const oSelected = oContext?.getObject();

          if (oSelected) {
            (this.byId("userSearchId") as Input).setValue(oSelected.Username);
          }

          // Filter
          this.applyFilters();

          oDialog.close();
        },
      });

      // Set data for List
      oList.setModel(oJsonModel);

      // Create Dialog
      const oDialog = new Dialog({
        title: "Select User",
        contentWidth: "400px",
        contentHeight: "500px",
        content: [oList],
        endButton: new Button({
          text: "Close",
          press: () => oDialog.close(),
        }),
      });

      oDialog.open();
    } catch (error) {
      MessageBox.error("Failed to load user search help.");
    }
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
}

/**
 * Called before the view is re-rendered.
 **/
// public  onBeforeRendering(): void {
//
//  }
/**
 * Called after the view has been rendered.
 **/
// public  onAfterRendering(): void {
//
//  }
/**
 * Called when the controller is destroyed.
 **/
// public onExit(): void {
//
//  }

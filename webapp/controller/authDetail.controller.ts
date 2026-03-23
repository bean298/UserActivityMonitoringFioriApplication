import Controller from "sap/fe/core/PageController";
import JSONModel from "sap/ui/model/json/JSONModel";
import ODataListBinding from "sap/ui/model/odata/v4/ODataListBinding";
import ODataModel from "sap/ui/model/odata/v4/ODataModel";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import Formatter from "useraudit/formatter/Formatter";
import MessageBox from "sap/m/MessageBox";
import MessageToast from "sap/m/MessageToast";
import Spreadsheet from "sap/ui/export/Spreadsheet";

export default class AuthDetail extends Controller {
  public formatter = Formatter;

  /**
   * Called when the controller is initialized.
   **/
  public onInit(): void {
    super.onInit();

    const oRouter = (this as any).getAppComponent().getRouter();
    if (oRouter) {
      oRouter
        .getRoute("AuthDetail")
        .attachPatternMatched(this._onObjectMatched, this);
    }
  }

  /**
   * Handles route matching for AuthDetail page
   * and loads detail data based on the navigation key.
   **/
  private async _onObjectMatched(oEvent: any): Promise<void> {
    // Get parameter from URL
    const sKey = oEvent.getParameter("arguments").key;

    const oView = this.getView();
    if (!oView || !sKey) return;

    oView.setBusy(true);

    try {
      const sSessionId = sKey.match(/'([^']+)'/)?.[1] || sKey;
      const oModel = (this as any).getAppComponent().getModel() as ODataModel;

      // Create a list binding to /UserAuthLog with $expand _Activity
      const oDetailBinding = oModel.bindList(
        "/UserAuthLog",
        undefined,
        undefined,
        [new Filter("SessionId", FilterOperator.EQ, sSessionId)],
        {
          $expand: "_Activity",
        },
      ) as ODataListBinding;

      // Executes the OData call
      const aContexts = await oDetailBinding.requestContexts(0, 1);

      // Set data into model
      if (aContexts.length > 0) {
        const oData = aContexts[0].getObject();
        const oDetailModel = new JSONModel(oData);

        oView.setModel(oDetailModel, "detailData");
      } else {
        MessageBox.error("Failed to load detail data. Please try again.");
      }
    } catch (oError) {
      MessageBox.error("Failed to load detail data. Please try again.");
    } finally {
      oView.setBusy(false);
    }
  }

  /**
   * Called when the user use filter status
   **/
  public onFilterActivityType(oEvent: any): void {
    const sKey = oEvent.getSource().getSelectedKey();

    const oTable = this.byId("ActivityTable") as any;
    const oBinding = oTable?.getBinding("rows");

    if (!oBinding) return;

    const aFilters: Filter[] = [];

    if (sKey) {
      aFilters.push(new Filter("ActivityType", FilterOperator.EQ, sKey));
    }

    oBinding.filter(aFilters);
  }

  /**
   * Exports the currently bound table data to an Excel file.
   * Uses sap.ui.export.Spreadsheet to generate the file
   * based on the current OData V4 list binding.
   */
  public onExportActivityExcel(): void {
    MessageBox.confirm("Do you want to export this data to Excel?", {
      title: "Confirm Export",
      actions: ["YES", "NO"],
      emphasizedAction: "YES",

      onClose: (oAction: string | null) => {
        if (oAction === "YES") {
          const oTable = this.byId("ActivityTable") as any;

          const aData = oTable.getModel("detailData").getProperty("/_Activity");

          const aCols = [
            { label: "Log ID", property: "LogId", width: 20 },
            { label: "Activity Type", property: "ActivityType", width: 15 },
            { label: "TCode", property: "TCode", width: 10 },
            { label: "TCode Name", property: "TCodeName", width: 20 },
            { label: "Message", property: "ActivityMessage", width: 40 },
            { label: "Time", property: "ActivityTime", width: 15 },
          ];

          const oSettings = {
            workbook: { columns: aCols },
            dataSource: aData,
            fileName: "ActivityLogs.xlsx",
            worker: false,
          };

          const oSheet = new Spreadsheet(oSettings);

          oSheet
            .build()
            .then(() => {
              MessageToast.show("Export successful!");
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

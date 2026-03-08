import Controller from "sap/fe/core/PageController";
import JSONModel from "sap/ui/model/json/JSONModel";
import ODataListBinding from "sap/ui/model/odata/v4/ODataListBinding";
import ODataModel from "sap/ui/model/odata/v4/ODataModel";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import Formatter from "useraudit/formatter/Formatter";
import MessageBox from "sap/m/MessageBox";

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
}

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

        oView.setModel(oUserDetailModel, "UserDetailData");
      } else {
        MessageBox.error("Failed to load user detail data. Please try again.");
      }
    } catch (oError) {
      MessageBox.error("Failed to load user detail data. Please try again.");
    } finally {
      oView.setBusy(false);
    }
  }
}

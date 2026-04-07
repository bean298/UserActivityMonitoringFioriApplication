import BaseComponent from "sap/fe/core/AppComponent";
import JSONModel from "sap/ui/model/json/JSONModel";

/**
 * @namespace useraudit
 */
export default class Component extends BaseComponent {
  public static metadata = {
    manifest: "json",
  };

  /**
   * The component is initialized by UI5 automatically during the startup of the app and calls the init method once.
   * @public
   * @override
   */
  public init(): void {
    super.init();

    // SET GLOBAL FILTER DATE
    const oToday = new Date();
    const oFrom = new Date();
    oFrom.setDate(oToday.getDate() - 3);

    const oGlobalModel = new JSONModel({
      fromDate: oFrom,
      toDate: oToday,
    });

    this.setModel(oGlobalModel, "global");
  }
}

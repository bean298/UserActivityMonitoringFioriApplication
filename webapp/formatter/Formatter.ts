import { ValueState } from "sap/ui/core/library";

const Formatter = {
  // Format Icon Login Result
  formatLoginResultIcon: function (sValue: string): string {
    if (sValue === "SUCCESS") {
      return "sap-icon://message-success";
    } else {
      return "sap-icon://message-error";
    }
  },

  //   Formate State Login Result
  formatLoginResultState: function (sValue: string): ValueState {
    if (sValue === "SUCCESS") {
      return ValueState.Success;
    } else {
      return ValueState.Error;
    }
  },

  //   Format State Login Message
  formatLoginMessageState: function (sValue: string): ValueState {
    if (sValue === "AU2" || sValue === "BU1" || sValue === "AUM") {
      return ValueState.Error;
    } else {
      return ValueState.Success;
    }
  },

  // Format Icon Login Message
  formatLoginMessageIcon: function (sValue: string): string {
    if (sValue === "AU2" || sValue === "BU1") {
      return "sap-icon://message-error";
    } else if (sValue === "AUM") {
      return "sap-icon://locked";
    } else {
      return "sap-icon://message-success";
    }
  },

  // Format Date
  formatISODate: function (oDate: any): string {
    if (!oDate) {
      return "00-00-0000";
    }

    const dateObj = new Date(oDate);

    // Check dateObj
    if (!isNaN(dateObj.getTime())) {
      const year = dateObj.getFullYear();
      const month = String(dateObj.getMonth() + 1).padStart(2, "0");
      const day = String(dateObj.getDate()).padStart(2, "0");

      return `${day}-${month}-${year}`;
    }

    return "00-00-0000";
  },

  // Format Time
  formatLogoutTime: function (oTime: any) {
    if (oTime === "12:00:00 AM") {
      return "00:00:00";
    }

    return oTime;
  },
};

export default Formatter;

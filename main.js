function Runner() {

    // Amazonのお届け日をGoogleカレンダーに記載
    console.log("/// ResistAmazonReceiptToCalendar Begin ///\n");
    ResistAmazonReceiptToCalendar();
    console.log("/// ResistAmazonReceiptToCalendar End ///\n");

    // Noshのお届け日をGoogleカレンダーに記載
    console.log("/// ResistNoshReceiptDateToCalendar Begin ///\n");
    ResistNoshReceiptToCalendar();
    console.log("/// ResistNoshReceiptDateToCalendar End ///\n");
}

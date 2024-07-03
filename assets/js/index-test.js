// Code js cho giao diện đẹp ấy mà

var current_fs, next_fs, previous_fs; //fieldsets
var left, opacity, scale; //fieldset properties which we will animate
var animating; //flag to prevent quick multi-click glitches
let isUploadCMSFile = false;
let isUploadFAPFile = false;

$(".next").click(function () {
  if (
    (!isUploadCMSFile || !isUploadFAPFile) &&
    $("fieldset").index($(this).parent()) == 0
  ) {
    alert("No file uploaded");
  } else {
    if ($("fieldset").index($(this).parent()) == 1) {
      let component = $("#exportedFile_name");
      let newFileName = component.val();
      if (newFileName == "" || newFileName == null) {
        $("#error_message").text("Filename cannot be blank");
      } else {
        newFileName = handleFilename(newFileName);
        $("#error_message").text("");
        exportToExcel(newFileName).then((result) => {
          if (result) {
            showNextFieldset(component);
          } else {
            alert("An error occurs while handling excel file(s)");
          }
        });
        resetFile();
      }
    }
    showNextFieldset($(this));
  }
});

$(".previous").click(function () {
  if ($("fieldset").index($(this).parent()) == 1) {
    let isReset = confirm(
      "Do you want to clear all uploaded data?\nAll uploaded data will be lost."
    );
    if (isReset) {
      resetData();
      showPreviousFieldset($(this));
    }
  } else {
    showPreviousFieldset($(this));
  }
});

function handleFilename(filename) {
  var re = /(?:\.([^.]+))?$/;
  var extension = re.exec(filename)[1];
  if (extension == "xlsx" || extension == "xls") {
    return filename;
  }
  return filename.trim() + ".xlsx";
}

function showNextFieldset(button) {
  if (animating) return false;
  animating = true;
  current_fs = button.parent();
  next_fs = button.parent().next();

  //activate next step on progressbar using the index of next_fs
  $("#progressbar li").eq($("fieldset").index(next_fs)).addClass("active");

  //show the next fieldset
  next_fs.show();
  //hide the current fieldset with style
  current_fs.animate(
    { opacity: 0 },
    {
      step: function (now, mx) {
        //as the opacity of current_fs reduces to 0 - stored in "now"
        //1. scale current_fs down to 80%
        scale = 1 - (1 - now) * 0.2;
        //2. bring next_fs from the right(50%)
        left = now * 50 + "%";
        //3. increase opacity of next_fs to 1 as it moves in
        opacity = 1 - now;
        current_fs.css({
          transform: "scale(" + scale + ")",
        });
        next_fs.css({ left: left, opacity: opacity });
      },
      duration: 800,
      complete: function () {
        current_fs.hide();
        animating = false;
      },
      //this comes from the custom easing plugin
      easing: "easeInOutBack",
    }
  );
}

function showPreviousFieldset(button) {
  if (animating) return false;
  resetFile();
  animating = true;

  current_fs = button.parent();
  previous_fs = button.parent().prev();

  //de-activate current step on progressbar
  $("#progressbar li")
    .eq($("fieldset").index(current_fs))
    .removeClass("active");

  //show the previous fieldset
  previous_fs.show();
  //hide the current fieldset with style
  current_fs.animate(
    { opacity: 0 },
    {
      step: function (now, mx) {
        //as the opacity of current_fs reduces to 0 - stored in "now"
        //1. scale previous_fs from 80% to 100%
        scale = 0.8 + (1 - now) * 0.2;
        //2. take current_fs to the right(50%) - from 0%
        left = (1 - now) * 50 + "%";
        //3. increase opacity of previous_fs to 1 as it moves in
        opacity = 1 - now;
        current_fs.css({ left: left });
        previous_fs.css({
          transform: "scale(" + scale + ")",
          opacity: opacity,
        });
      },
      duration: 800,
      complete: function () {
        current_fs.hide();
        animating = false;
      },
      //this comes from the custom easing plugin
      easing: "easeInOutBack",
    }
  );
}

function showSpinner() {
  $(".spinner").show();
  $("#msform").hide();
}

function hideSpinner() {
  $(".spinner").hide();
  $("#msform").show();
}

// Hết code js giao diện rồi

// Code xử lý file excel
let markData = [];
let headers = [];

// score constant
const SCORE_FAIL_OR_REFER = 1;
const SCORE_PASS = 65;
const SCORE_MERIT = 80;
const SCORE_DISTINCTION = 100;
const INVALID_ASM = -1;
const ERROR_ASM = -2;

// file type
const CMS_TYPE = 1;
const FAP_TYPE = 2;

// score constant by word
const SCORE_FAIL_OR_REFER_W = ["fail", "refer"];
const SCORE_PASS_W = "pass";
const SCORE_MERIT_W = "merit";
const SCORE_DISTINCTION_W = "distinction";
const SCORE_INVALID_W = "Lỗi điểm second chance";
const SCORE_ERROR_W = "Error";

const SCORE_ERROR = "-";

// min and max score of each mark level

const SCORE_FAIL_OR_REFER_RANGE = [1];
const SCORE_PASS_RANGE = [65, 79];
const SCORE_MERIT_RANGE = [80, 89];
const SCORE_DISTINCTION_RANGE = [90, 100];
const SCORE_ERROR_ASM_RANGE = [-1];
const SCORE_INVALID_ASM_RANGE = [-2];

// data
let exportedData = [];
let exportedSheetName = [];

// Header of CMS and Final excel files
let cms_header_arr = [];
let final_header_arr = [];

// Array of score column indexes in CMS excel file
const CMS_SCORE_COLUMNS = ["1", "2", "3", "4"];

// Email column in FAI excel file
const FAI_EMAIL_COLUMN = 4;

// Number of columns in combined files
const MAX_COLUMNS_NUMBER = 6;

// open file dialog event listener for button upload
// $("#upload").on("click", function () {
//   file.click();
// });
$("#file_fap").on("change", function () {
  isUploadFAPFile = true;
  let button = $(this);
  const files = this.files;
  Object.keys(files).forEach(function (i) {
    if (files[i]) {
      readFile(files[i], FAP_TYPE);
    }
  });
});

$("#file_cms").on("change", function () {
  isUploadCMSFile = true;
  let button = $(this);
  const files = this.files;
  Object.keys(files).forEach(function (i) {
    if (files[i]) {
      readFile(files[i], CMS_TYPE);
    }
  });
});

var delay = (function () {
  var timer = 0;
  return function (callback, ms) {
    clearTimeout(timer);
    timer = setTimeout(callback, ms);
  };
})();

function readFile(file, file_type) {
  // initialize new Excel Reader SheetJS
  let reader = new FileReader();
  // handle event for reader variable when upload file
  reader.onload = function (e) {
    // get data from input
    let data = e.target.result;
    // read data
    let workbook = XLSX.read(data, {
      type: "binary",
    });
    //
    let col_count = 0;
    const sheet_count = workbook.SheetNames.length;
    // get first sheet name from excel file
    let sheet_name = workbook.SheetNames[0];
    // get data from sheet (Convert to json)
    let sheet = workbook.Sheets[sheet_name];
    // Set header = 1 to get header of excel sheet
    // get header of sheet columns
    let jsonHeader = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: null,
    })[0];
    // get data of sheet only into key => value array
    let jsonData = XLSX.utils.sheet_to_json(sheet, { defval: null });
    // set up col count to distinguish FAP file or CMS file
    col_count = Object.keys(jsonData[0]).length;
    headers = jsonData.length > 0 ? jsonData[0] : [];
    if (file_type == CMS_TYPE) {
      cms_header_arr = jsonHeader;
      markData = jsonData;
    } else {
      final_header_arr = jsonHeader;
      exportedData = jsonData;
    }
  };
  reader.readAsBinaryString(file);
}

// Function to reset input
function resetFile() {
  let files = document.querySelectorAll(".file");
  files.forEach(function (file) {
    file.value = "";
  });
  isUploadFile = false;
}

// Reset all data
function resetData() {
  resetFile();
  $("#exportedFile_name").val("");
  markData = [];
  headers = [];
  exportedData = [];
  exportedSheetName = [];
  cms_header_arr = [];
  final_header_arr = [];
}

function getFinalScore(final_score, second_chance_score) {
  if (final_score < SCORE_PASS_RANGE[0]) {
    if (second_chance_score == SCORE_ERROR) {
      second_chance_score = 1;
    }
    if (second_chance_score > SCORE_PASS_RANGE[1]) {
      return SCORE_ERROR_ASM_RANGE[0];
    }
    return second_chance_score;
  }
  if (second_chance_score != SCORE_ERROR)
  {
    return SCORE_ERROR_ASM_RANGE[0];
  }
  return final_score;
}

// Function to handle score of asm 1 and asm 2
// Return the score by number
function handleScore(data) {
  let final_score = data[cms_header_arr[CMS_SCORE_COLUMNS[2]]];
  let second_chance_score = data[cms_header_arr[CMS_SCORE_COLUMNS[3]]];
  let temp = getFinalScore(final_score, second_chance_score);
  return getFinalScore(final_score, second_chance_score);
}

// Function to write data to Final columns in exported excel file
function handleData() {
  // Map value by email from CMS to FAP File
  showSpinner();
  markData.forEach(function (item) {
    let value = exportedData.find(
      (obj) =>
        obj[final_header_arr[4]].toLowerCase().trim() ===
        item[cms_header_arr[0]].toLowerCase().trim()
    );
    // if the value is mapped
    if (value) {
      // Check if student is not fail in attendance
      if (exportedData[exportedData.indexOf(value)][final_header_arr[5]] != 0) {
        // Get score
        let score = handleScore(item);
        // Check valid of asm 1st chance and 2nd chance
        switch (score) {
          case INVALID_ASM: {
            exportedData[exportedData.indexOf(value)][final_header_arr[5]] =
              SCORE_INVALID_W;
            break;
          }
          case ERROR_ASM: {
            exportedData[exportedData.indexOf(value)][final_header_arr[5]] =
              SCORE_ERROR_W;
            break;
          }
          default: {
            exportedData[exportedData.indexOf(value)][final_header_arr[5]] =
              score / 10;
            break;
          }
        }
      } else {
        exportedData[exportedData.indexOf(value)][final_header_arr[5]] = 0;
      }
    }
  });
  hideSpinner();
}

// Export excel function
async function exportToExcel(fileName) {
  let isConnect = false;
  try {
    const XLSX = await import(
      "https://cdn.sheetjs.com/xlsx-0.19.2/package/xlsx.mjs"
    );
    // Create a new workbook and add a worksheet
    handleData();
    var wb = XLSX.utils.book_new();
    // Create a new workbook
    // Convert the array to a worksheet
    var ws = XLSX.utils.json_to_sheet(exportedData);

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, "Exported Data");
    // Save the workbook to an Excel file
    XLSX.writeFile(wb, fileName || "data.xlsx");
    isConnect = true;
  } catch (error) {
    isConnect = false;
  }
  return isConnect;
}

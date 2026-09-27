/*******************************************************
 * 1) CONFIG
 *******************************************************/

const SYSTEM_CONFIG = {

  // Spreadsheet الجديد - الحضور
  ATTENDANCE_SPREADSHEET_ID:
    "1k55FqxrS5qhP3zoASt7ia2TWMv8cFn3XaYRrFKFN8nE",

  // Spreadsheet القديم - الشهادات
  LEGACY_SPREADSHEET_ID:
    "1mQRoDtOXX3WN1YLUJJUhMwn5eZV7nBh_yPXQpjxL1Qg",

  // اسم Sheet الكورسات القديمة
  COURSES_SHEET:
    "Courses",

  // اسم Sheet الشهادات القديمة
  CERTIFICATES_SHEET:
    "Certificates",

  // Folder الرئيسي للشهادات
  CERTIFICATES_PARENT_FOLDER_ID:
    "1OhY-33kDLWFDQ7L7BmMBSB0vuCBJE1jZ",

  // أسماء أعمدة الحضور
  HEADERS: {
    ID: "ID",
    NAME: "الاسم",
    COUNT: "عدد مرات الحضور",
    CERTIFICATE: "الشهادة"
  },

  // Prefixes للـProperties
  PROPERTY: {
    COURSE_SHEET: "COURSE_SHEET_",
    MEETING_SHEET: "MEETING_SHEET_",
    ACTIVITY_CREATED_AT: "ACTIVITY_CREATED_AT_"
  }
};


/*******************************************************
 * 2) BASIC HELPERS
 *******************************************************/

/**
 * الحصول على Spreadsheet الحضور
 */
function getAttendanceSpreadsheet_() {

  return SpreadsheetApp.openById(
    SYSTEM_CONFIG.ATTENDANCE_SPREADSHEET_ID
  );
}


/**
 * الحصول على Spreadsheet القديم
 */
function getLegacySpreadsheet_() {

  return SpreadsheetApp.openById(
    SYSTEM_CONFIG.LEGACY_SPREADSHEET_ID
  );
}


/**
 * الحصول على Courses القديمة
 */
function getCoursesSheet_() {

  const ss = getLegacySpreadsheet_();

  const sheet = ss.getSheetByName(
    SYSTEM_CONFIG.COURSES_SHEET
  );

  if (!sheet) {
    throw new Error(
      'لم يتم العثور على Sheet باسم "Courses"'
    );
  }

  return sheet;
}


/**
 * الحصول على Certificates القديمة
 */
function getLegacyCertificatesSheet_() {

  const ss = getLegacySpreadsheet_();

  return ss.getSheetByName(
    SYSTEM_CONFIG.CERTIFICATES_SHEET
  );
}


/**
 * الحصول على Folder الشهادات الرئيسي
 */
function getCertificatesParentFolder_() {

  return DriveApp.getFolderById(
    SYSTEM_CONFIG.CERTIFICATES_PARENT_FOLDER_ID
  );
}


/**
 * تنظيف اسم
 */
function cleanName_(value) {

  return String(value || "")
    .trim()
    .replace(/\s+/g, " ");
}


/*******************************************************
 * 3) PROPERTIES
 *******************************************************/

/**
 * حفظ Sheet ID الخاص بالكورس
 */
function saveCourseSheetId_(courseId, sheetId) {

  PropertiesService
    .getScriptProperties()
    .setProperty(
      SYSTEM_CONFIG.PROPERTY.COURSE_SHEET + courseId,
      String(sheetId)
    );
}


/**
 * الحصول على Sheet ID الخاص بالكورس
 */
function getCourseSheetId_(courseId) {

  return PropertiesService
    .getScriptProperties()
    .getProperty(
      SYSTEM_CONFIG.PROPERTY.COURSE_SHEET + courseId
    );
}


/**
 * حفظ أن Sheet اجتماع
 */
function saveMeetingSheet_(sheetId) {

  PropertiesService
    .getScriptProperties()
    .setProperty(
      SYSTEM_CONFIG.PROPERTY.MEETING_SHEET + sheetId,
      "true"
    );
}


/**
 * هل الـSheet اجتماع؟
 */
function isMeetingSheet_(sheetId) {

  const value = PropertiesService
    .getScriptProperties()
    .getProperty(
      SYSTEM_CONFIG.PROPERTY.MEETING_SHEET + sheetId
    );

  return value === "true";
}


/**
 * حفظ تاريخ إنشاء النشاط
 */
function saveActivityCreatedAt_(sheetId) {

  PropertiesService
    .getScriptProperties()
    .setProperty(
      SYSTEM_CONFIG.PROPERTY.ACTIVITY_CREATED_AT + sheetId,
      new Date().toISOString()
    );
}


/**
 * الحصول على تاريخ إنشاء النشاط
 */
function getActivityCreatedAt_(sheetId) {

  return PropertiesService
    .getScriptProperties()
    .getProperty(
      SYSTEM_CONFIG.PROPERTY.ACTIVITY_CREATED_AT + sheetId
    ) || "";
}


/**
 * تنسيق التاريخ
 */
function formatActivityDate_(isoDate) {

  if (!isoDate) {
    return "";
  }

  try {

    return Utilities.formatDate(
      new Date(isoDate),
      getAttendanceSpreadsheet_().getSpreadsheetTimeZone(),
      "dd/MM/yyyy"
    );

  } catch (err) {

    return "";
  }
}


/*******************************************************
 * 4) GENERAL SHEET HELPERS
 *******************************************************/

/**
 * الحصول على Sheet نشاط
 */
function getActivitySheetById_(sheetId) {

  const ss = getAttendanceSpreadsheet_();

  const sheet = ss.getSheets()
    .find(function (s) {

      return String(s.getSheetId()) === String(sheetId);

    });

  if (!sheet) {

    throw new Error(
      "لم يتم العثور على Sheet النشاط."
    );
  }

  return sheet;
}


/**
 * إنشاء Headers الخاصة بالنشاط
 */
function createActivitySheet_(sheet, activityName) {

  const headers = [
    "ID",
    "الاسم",
    "عدد مرات الحضور",
    "الشهادة"
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold");

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 300);

  sheet
    .getRange("C:C")
    .setNumberFormat("0");

  sheet
    .getRange("A:A")
    .setNumberFormat("@");

  sheet
    .getRange("B:B")
    .setNumberFormat("@");

  /*
   * مهم جدًا: كل أعمدة المحاضرات (من العمود E فصاعدًا) لازم تكون
   * "نص" (Plain Text) من البداية، وليس تاريخ. لو سبتها بدون تنسيق،
   * Google Sheets هيفسّر أي قيمة تاريخ/وقت نكتبها فيها كـ Date حقيقي،
   * وساعتها لما نقرأها تانى من Apps Script هترجع كـ JS Date Object
   * وتظهر فى الواجهة بشكل خام غريب زي:
   * "Wed Sep 23 2026 00:40:00 GMT+0300 (...)"
   * بدل الشكل المُنسّق اللي احنا بنحطه فعليًا.
   */
  sheet
    .getRange("E:ZZ")
    .setNumberFormat("@");

  return sheet;
}


/**
 * آخر صف مستخدم
 */
function getLastDataRow_(sheet) {

  const lastRow = sheet.getLastRow();

  return Math.max(lastRow, 1);
}


/*******************************************************
 * 5) COURSE ID
 *******************************************************/

/**
 * الحصول على CourseID جديد
 *
 * CourseID يتم أخذه من Courses القديمة
 */
function getNextCourseId_() {

  const sheet = getCoursesSheet_();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 1;
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues();

  let maxId = 0;

  values.forEach(function (row) {

    const value = Number(row[0]);

    if (!isNaN(value) && value > maxId) {
      maxId = value;
    }

  });

  return maxId + 1;
}


/*******************************************************
 * 6) CREATE COURSE
 *******************************************************/

/**
 * إنشاء كورس جديد
 */
function createCourse(courseName) {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    courseName = cleanName_(courseName);

    if (!courseName) {

      throw new Error(
        "يجب إدخال اسم الكورس."
      );
    }


    const coursesSheet = getCoursesSheet_();

    const lastRow = coursesSheet.getLastRow();


    // منع تكرار اسم الكورس
    if (lastRow >= 2) {

      const existingNames = coursesSheet
        .getRange(2, 2, lastRow - 1, 1)
        .getValues();

      const exists = existingNames.some(function (row) {

        return cleanName_(row[0]) === courseName;

      });

      if (exists) {

        throw new Error(
          "هذا الكورس موجود بالفعل."
        );
      }
    }


    const courseId = getNextCourseId_();

    const attendanceSS =
      getAttendanceSpreadsheet_();


    // إنشاء Sheet
    const sheet =
      attendanceSS.insertSheet(courseName);


    createActivitySheet_(
      sheet,
      courseName
    );


    // حفظ تاريخ الإنشاء
    saveActivityCreatedAt_(
      sheet.getSheetId()
    );


    // إنشاء Folder
    const parentFolder =
      getCertificatesParentFolder_();

    const courseFolder =
      parentFolder.createFolder(courseName);


    // حفظ Sheet ID
    saveCourseSheetId_(
      courseId,
      sheet.getSheetId()
    );


    // إضافة للكورسات القديمة
    coursesSheet.appendRow([
      courseId,
      courseName,
      courseFolder.getId(),
      courseFolder.getUrl()
    ]);


    return {

      success: true,

      type: "course",

      courseId: courseId,

      name: courseName,

      sheetId: sheet.getSheetId(),

      folderId: courseFolder.getId(),

      folderUrl: courseFolder.getUrl(),

      createdAt:
        getActivityCreatedAt_(
          sheet.getSheetId()
        ),

      createdDate:
        formatActivityDate_(
          getActivityCreatedAt_(
            sheet.getSheetId()
          )
        )
    };


  } finally {

    lock.releaseLock();

  }
}


/*******************************************************
 * 7) CREATE MEETING
 *******************************************************/

/**
 * إنشاء اجتماع جديد
 *
 * الاجتماع:
 * - لا يدخل Courses
 * - لا ينشئ Folder
 * - لا يعالج شهادات
 */
function createMeeting(meetingName) {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    meetingName = cleanName_(meetingName);

    if (!meetingName) {

      throw new Error(
        "يجب إدخال اسم الاجتماع."
      );
    }


    const ss =
      getAttendanceSpreadsheet_();


    // منع Sheet بنفس الاسم
    if (ss.getSheetByName(meetingName)) {

      throw new Error(
        "يوجد نشاط بنفس الاسم بالفعل."
      );
    }


    const sheet =
      ss.insertSheet(meetingName);


    createActivitySheet_(
      sheet,
      meetingName
    );


    saveMeetingSheet_(
      sheet.getSheetId()
    );


    saveActivityCreatedAt_(
      sheet.getSheetId()
    );


    return {

      success: true,

      type: "meeting",

      name: meetingName,

      sheetId: sheet.getSheetId(),

      createdAt:
        getActivityCreatedAt_(
          sheet.getSheetId()
        ),

      createdDate:
        formatActivityDate_(
          getActivityCreatedAt_(
            sheet.getSheetId()
          )
        )
    };


  } finally {

    lock.releaseLock();

  }
}


/*******************************************************
 * 8) FIND COURSE INFORMATION
 *******************************************************/

/**
 * الحصول على معلومات CourseID
 */
function getCourseInfoById_(courseId) {

  const sheet = getCoursesSheet_();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }


  const values = sheet
    .getRange(2, 1, lastRow - 1, 4)
    .getValues();


  for (let i = 0; i < values.length; i++) {

    if (
      String(values[i][0]) ===
      String(courseId)
    ) {

      return {

        courseId: values[i][0],

        name: cleanName_(values[i][1]),

        folderId:
          String(values[i][2] || ""),

        folderUrl:
          String(values[i][3] || "")

      };
    }
  }

  return null;
}


/**
 * الحصول على CourseID من Sheet ID
 */
function getCourseIdBySheetId_(sheetId) {

  const props =
    PropertiesService.getScriptProperties();

  const allProps =
    props.getProperties();


  const prefix =
    SYSTEM_CONFIG.PROPERTY.COURSE_SHEET;


  for (const key in allProps) {

    if (key.indexOf(prefix) !== 0) {
      continue;
    }


    if (
      String(allProps[key]) ===
      String(sheetId)
    ) {

      return key.substring(
        prefix.length
      );
    }
  }


  return null;
}


/*******************************************************
 * 9) ACTIVITY LIST
 *******************************************************/

/**
 * الحصول على كل الأنشطة
 *
 * Courses + Meetings
 */
function getActivities() {

  const result = [];


  /**********************
   * COURSES
   **********************/

  const coursesSheet =
    getCoursesSheet_();

  const lastRow =
    coursesSheet.getLastRow();


  if (lastRow >= 2) {

    const rows =
      coursesSheet
        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )
        .getValues();


    rows.forEach(function (row) {

      const courseId =
        row[0];

      const courseName =
        cleanName_(row[1]);


      if (!courseId || !courseName) {
        return;
      }


      const sheetId =
        getCourseSheetId_(
          courseId
        );


      if (!sheetId) {
        return;
      }


      let sheet = null;


      try {

        sheet =
          getActivitySheetById_(
            sheetId
          );

      } catch (err) {

        // Sheet اتحذفت يدويًا
        return;
      }


      const summary =
        getActivitySummary_(
          sheet
        );


      result.push({

        type: "course",

        courseId:
          courseId,

        name:
          courseName,

        sheetId:
          sheet.getSheetId(),

        folderId:
          String(row[2] || ""),

        folderUrl:
          String(row[3] || ""),

        createdAt:
          getActivityCreatedAt_(
            sheet.getSheetId()
          ),

        createdDate:
          formatActivityDate_(
            getActivityCreatedAt_(
              sheet.getSheetId()
            )
          ),

        peopleCount:
          summary.peopleCount,

        actualAttendees:
          summary.actualAttendees,

        lectures:
          summary.lectures
      });

    });
  }


  /**********************
   * MEETINGS
   **********************/

  const ss =
    getAttendanceSpreadsheet_();

  const sheets =
    ss.getSheets();


  sheets.forEach(function (sheet) {

    const sheetId =
      sheet.getSheetId();


    if (!isMeetingSheet_(sheetId)) {
      return;
    }


    // لو الـSheet تم ربطه لاحقًا بكورس
    if (getCourseIdBySheetId_(sheetId)) {
      return;
    }


    const summary =
      getActivitySummary_(
        sheet
      );


    result.push({

      type: "meeting",

      courseId: null,

      name:
        sheet.getName(),

      sheetId:
        sheetId,

      folderId: null,

      folderUrl: null,

      createdAt:
        getActivityCreatedAt_(
          sheetId
        ),

      createdDate:
        formatActivityDate_(
          getActivityCreatedAt_(
            sheetId
          )
        ),

      peopleCount:
        summary.peopleCount,

      actualAttendees:
        summary.actualAttendees,

      lectures:
        summary.lectures
    });

  });


  return result;
}


/**
 * Summary لنشاط
 */
function getActivitySummary_(sheet) {

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();


  let peopleCount = 0;
  let actualAttendees = 0;


  if (lastRow >= 2) {

    const values =
      sheet
        .getRange(
          2,
          2,
          lastRow - 1,
          2
        )
        .getValues();


    values.forEach(function (row) {

      const name =
        cleanName_(row[0]);

      const count =
        Number(row[1]) || 0;


      // أي اسم موجود في B
      // يدخل في عدد الأشخاص
      if (name) {
        peopleCount++;
      }


      // حضور فعلي
      if (name && count > 0) {
        actualAttendees++;
      }

    });
  }


  const lectures =
    getLectures(
      sheet.getSheetId()
    );


  lectures.forEach(function (lecture) {

    lecture.attendanceCount =
      countLectureAttendance_(
        sheet,
        lecture.column
      );

  });


  return {

    peopleCount:
      peopleCount,

    actualAttendees:
      actualAttendees,

    lectures:
      lectures
  };
}


/**
 * عدد حضور محاضرة
 */
function countLectureAttendance_(
  sheet,
  columnNumber
) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return 0;
  }


  const values =
    sheet
      .getRange(
        2,
        columnNumber,
        lastRow - 1,
        1
      )
      .getValues();


  let count = 0;


  values.forEach(function (row) {

    if (
      row[0] !== "" &&
      row[0] !== null
    ) {
      count++;
    }

  });


  return count;
}


/*******************************************************
 * 10) LECTURES
 *******************************************************/

/**
 * التأكد من عدم تكرار اسم المحاضرة
 */
function lectureNameExists_(
  sheet,
  lectureName
) {

  lectureName =
    cleanName_(lectureName);


  const lastColumn =
    sheet.getLastColumn();


  if (lastColumn < 5) {
    return false;
  }


  const headers =
    sheet
      .getRange(
        1,
        5,
        1,
        lastColumn - 4
      )
      .getValues()[0];


  return headers.some(function (value) {

    return (
      cleanName_(value) ===
      lectureName
    );

  });
}


/**
 * إضافة محاضرة
 */
function addLecture(
  sheetId,
  lectureName
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const sheet =
      getActivitySheetById_(
        sheetId
      );


    lectureName =
      cleanName_(lectureName);


    if (!lectureName) {

      throw new Error(
        "يجب إدخال اسم المحاضرة."
      );
    }


    if (
      lectureNameExists_(
        sheet,
        lectureName
      )
    ) {

      throw new Error(
        "هذه المحاضرة موجودة بالفعل."
      );
    }


    const newColumn =
      Math.max(
        5,
        sheet.getLastColumn() + 1
      );


    sheet
      .getRange(
        1,
        newColumn
      )
      .setValue(
        lectureName
      );


    sheet
      .getRange(
        1,
        newColumn
      )
      .setFontWeight(
        "bold"
      );


    sheet.setColumnWidth(
      newColumn,
      180
    );


    /*
     * نفس سبب التنسيق فى createActivitySheet_: نضمن إن عمود
     * المحاضرة الجديد كله (مش بس رأسه) نصّي من البداية، حتى لو
     * الشيت كبر لآلاف الصفوف لاحقًا.
     */
    sheet
      .getRange(
        1,
        newColumn,
        Math.max(sheet.getMaxRows(), 1000),
        1
      )
      .setNumberFormat("@");


    return {

      success: true,

      column:
        newColumn,

      name:
        lectureName

    };


  } finally {

    lock.releaseLock();

  }
}


/**
 * الحصول على المحاضرات
 */
function getLectures(sheetId) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  const lastColumn =
    sheet.getLastColumn();


  if (lastColumn < 5) {
    return [];
  }


  const headers =
    sheet
      .getRange(
        1,
        5,
        1,
        lastColumn - 4
      )
      .getValues()[0];


  return headers
    .map(function (name, index) {

      if (!cleanName_(name)) {
        return null;
      }


      return {

        name:
          cleanName_(name),

        column:
          index + 5,

        attendanceCount:
          countLectureAttendance_(
            sheet,
            index + 5
          )
      };

    })
    .filter(function (item) {

      return item !== null;

    });
}


/**
 * تعديل اسم المحاضرة
 */
function renameLecture(
  sheetId,
  columnNumber,
  newLectureName
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  newLectureName =
    cleanName_(
      newLectureName
    );


  if (!newLectureName) {

    throw new Error(
      "يجب إدخال اسم المحاضرة."
    );
  }


  if (columnNumber < 5) {

    throw new Error(
      "رقم المحاضرة غير صحيح."
    );
  }


  const oldName =
    cleanName_(
      sheet
        .getRange(
          1,
          columnNumber
        )
        .getValue()
    );


  if (
    oldName !== newLectureName &&
    lectureNameExists_(
      sheet,
      newLectureName
    )
  ) {

    throw new Error(
      "يوجد محاضرة أخرى بنفس الاسم."
    );
  }


  sheet
    .getRange(
      1,
      columnNumber
    )
    .setValue(
      newLectureName
    );


  return {

    success: true,

    oldName:
      oldName,

    newName:
      newLectureName,

    column:
      columnNumber
  };
}


/*******************************************************
 * 11) QR PARSING
 *******************************************************/

/**
 * قراءة QR
 *
 * الشكل:
 *
 * AB001--توماس ماجد فطيم عطية
 */
function parseQrData_(qrData) {

  qrData =
    String(qrData || "")
      .trim();


  if (!qrData) {

    throw new Error(
      "QR Code فارغ."
    );
  }


  const separator =
    "--";


  const index =
    qrData.indexOf(separator);


  if (index === -1) {

    throw new Error(
      "صيغة QR غير صحيحة."
    );
  }


  const id =
    qrData
      .substring(
        0,
        index
      )
      .trim();


  const name =
    qrData
      .substring(
        index + separator.length
      )
      .trim();


  if (!id || !name) {

    throw new Error(
      "بيانات QR غير مكتملة."
    );
  }


  return {

    id: id,

    name: cleanName_(name)

  };
}


/*******************************************************
 * 12) ATTENDANCE HELPERS
 *******************************************************/

/**
 * البحث عن ID
 */
function findRowById_(
  sheet,
  personId
) {

  personId =
    String(personId || "")
      .trim();


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return null;
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  for (let i = 0; i < values.length; i++) {

    if (
      String(values[i][0])
        .trim() === personId
    ) {

      return i + 2;
    }
  }


  return null;
}


/**
 * تنسيق أي تاريخ/وقت بنفس صيغة الحضور المستخدمة فى كل الشيت
 * (مفصولة عن getAttendanceTimestamp_ حتى تقدر المزامنة الأوفلاين
 * تنسّق وقت التسجيل الحقيقي وقت الحدث، مش وقت المزامنة)
 *
 * الصيغة: hh:mm:ss ص/م  yyyy/MM/dd
 * مثال:   07:32:15 ص  2026/09/23
 *
 * ملاحظة: القيمة دايمًا بترجع كنص (String) وليس تاريخ، والعمود نفسه
 * مضبوط كـ "نص" (setNumberFormat("@")) فى createActivitySheet_ /
 * addLecture، عشان Google Sheets ميحولهاش لـ Date تلقائيًا، وده اللي
 * كان بيسبب ظهور المنطقة الزمنية (GMT+0300) وشكل التاريخ الغريب.
 */
function formatAttendanceTimestamp_(date) {

  const timezone =
    getAttendanceSpreadsheet_()
      .getSpreadsheetTimeZone();


  return Utilities.formatDate(
    date,
    timezone,
    "hh:mm:ss a yyyy/MM/dd"
  )
  .replace("AM", "ص")
  .replace("PM", "م");
}


/**
 * Timestamp — وقت الآن
 */
function getAttendanceTimestamp_() {

  return formatAttendanceTimestamp_(
    new Date()
  );
}


/**
 * تطبيع أي قيمة قادمة من خلية محاضرة قبل إرجاعها للواجهة.
 *
 * لأسباب تاريخية (بيانات قديمة قبل تفعيل setNumberFormat("@"))،
 * بعض الخلايا ممكن تكون لسه متخزنة كـ Date حقيقي جوه Sheets بدل
 * نص. الدالة دي بتتأكد إن أي قيمة من النوع ده بترجع كنص منسّق
 * بنفس شكل باقي التوقيتات، بدل ما ترجع Date Object خام للواجهة.
 */
function normalizeTimestampValue_(value) {

  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {

    return formatAttendanceTimestamp_(value);
  }

  return value;
}


/**
 * تسجيل صف جديد
 */
function createPersonRow_(
  sheet,
  personId,
  personName
) {

  const newRow =
    sheet.getLastRow() + 1;


  sheet
    .getRange(
      newRow,
      1,
      1,
      4
    )
    .setValues([[
      personId || "",
      cleanName_(personName),
      0,
      ""
    ]]);


  return newRow;
}


/*******************************************************
 * 13) QR ATTENDANCE
 *******************************************************/

/**
 * تسجيل حضور عن طريق QR
 *
 * QR يتم تسجيله مباشرة
 */
function recordQrAttendance(
  sheetId,
  lectureColumn,
  qrData
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const sheet =
      getActivitySheetById_(
        sheetId
      );


    lectureColumn =
      Number(lectureColumn);


    if (lectureColumn < 5) {

      throw new Error(
        "المحاضرة غير صحيحة."
      );
    }


    const person =
      parseQrData_(
        qrData
      );


    let rowNumber =
      findRowById_(
        sheet,
        person.id
      );


    /****************************************
     * الشخص موجود
     ****************************************/

    if (rowNumber) {

      const currentName =
        cleanName_(
          sheet
            .getRange(
              rowNumber,
              2
            )
            .getValue()
        );


      const actualName =
        currentName ||
        person.name;


      const attendanceCell =
        sheet
          .getRange(
            rowNumber,
            lectureColumn
          );


      const already =
        attendanceCell
          .getValue();


      // نفس الشخص في نفس المحاضرة
      if (
        already !== "" &&
        already !== null
      ) {

        return {

          success: false,

          alreadyAttended: true,

          errorCode:
            "ALREADY_ATTENDED",

          message:
            "تم إضافة " +
            actualName +
            " بالفعل",

          id:
            person.id,

          name:
            actualName,

          count:
            Number(
              sheet
                .getRange(
                  rowNumber,
                  3
                )
                .getValue()
            ) || 0,

          lectureColumn:
            lectureColumn,

          timestamp:
            normalizeTimestampValue_(already)

        };
      }


      // الاسم الموجود في Sheet هو المعتمد
      if (!currentName) {

        sheet
          .getRange(
            rowNumber,
            2
          )
          .setValue(
            person.name
          );
      }


      // تسجيل Timestamp
      const timestamp =
        getAttendanceTimestamp_();


      attendanceCell
        .setValue(
          timestamp
        );


      // زيادة العدد
      const countCell =
        sheet
          .getRange(
            rowNumber,
            3
          );


      const oldCount =
        Number(
          countCell.getValue()
        ) || 0;


      const newCount =
        oldCount + 1;


      countCell
        .setValue(
          newCount
        );


      return {

        success: true,

        alreadyAttended: false,

        id:
          person.id,

        name:
          actualName || person.name,

        count:
          newCount,

        lectureColumn:
          lectureColumn,

        timestamp:
          timestamp
      };
    }


    /****************************************
     * الشخص غير موجود
     ****************************************/

    rowNumber =
      createPersonRow_(
        sheet,
        person.id,
        person.name
      );


    const timestamp =
      getAttendanceTimestamp_();


    sheet
      .getRange(
        rowNumber,
        lectureColumn
      )
      .setValue(
        timestamp
      );


    sheet
      .getRange(
        rowNumber,
        3
      )
      .setValue(
        1
      );


    return {

      success: true,

      alreadyAttended: false,

      id:
        person.id,

      name:
        person.name,

      count:
        1,

      lectureColumn:
        lectureColumn,

      timestamp:
        timestamp
    };


  } finally {

    lock.releaseLock();

  }
}


/*******************************************************
 * 14) NAME NORMALIZATION
 *******************************************************/

/**
 * Normalize عربي للبحث
 */
function normalizeArabicName_(name) {

  return String(name || "")
    .toLowerCase()
    .trim()

    // إزالة التشكيل
    .replace(/[\u064B-\u065F\u0670]/g, "")

    // أ إ آ
    .replace(/[أإآ]/g, "ا")

    // ى -> ي
    .replace(/ى/g, "ي")

    // ة -> ه
    .replace(/ة/g, "ه")

    // ؤ
    .replace(/ؤ/g, "و")

    // ئ
    .replace(/ئ/g, "ي")

    // چ
    .replace(/چ/g, "ج")

    // ڤ
    .replace(/ڤ/g, "ف")

    // گ
    .replace(/گ/g, "ك")

    // مسافات
    .replace(/\s+/g, " ");
}


/*******************************************************
 * 15) SEARCH MANUAL ATTENDANCE
 *******************************************************/

/**
 * البحث الجزئي بالاسم
 *
 * يستخدم من شاشة Scan
 */
function searchManualAttendancePeople(
  sheetId,
  lectureColumn,
  query
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  query =
    normalizeArabicName_(
      query
    );


  if (!query) {
    return [];
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const lastColumn =
    sheet.getLastColumn();


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        Math.max(
          4,
          lastColumn
        )
      )
      .getValues();


  const results = [];


  values.forEach(function (row, index) {

    const id =
      String(row[0] || "")
        .trim();


    const name =
      cleanName_(row[1]);


    const count =
      Number(row[2]) || 0;


    if (!name) {
      return;
    }


    const normalizedName =
      normalizeArabicName_(
        name
      );


    if (
      normalizedName.indexOf(query) === -1
    ) {
      return;
    }


    let alreadyAttended = false;
    let timestamp = "";


    if (
      lectureColumn &&
      Number(lectureColumn) >= 5 &&
      Number(lectureColumn) <= lastColumn
    ) {

      const value =
        row[
          Number(lectureColumn) - 1
        ];


      if (
        value !== "" &&
        value !== null
      ) {

        alreadyAttended = true;
        timestamp = normalizeTimestampValue_(value);
      }
    }


    results.push({

      rowNumber:
        index + 2,

      id:
        id,

      name:
        name,

      count:
        count,

      alreadyAttended:
        alreadyAttended,

      timestamp:
        timestamp,

      hasId:
        !!id

    });

  });


  // ترتيب أبجدي
  results.sort(function (a, b) {

    return a.name.localeCompare(
      b.name,
      "ar"
    );

  });


  return results;
}


/*******************************************************
 * 16) MANUAL ATTENDANCE
 *******************************************************/

/**
 * تسجيل حضور يدوي
 *
 * rowNumber اختياري
 *
 * لو الاسم مكرر:
 * requiresSelection = true
 */
function recordManualAttendance(
  sheetId,
  lectureColumn,
  personName,
  rowNumber
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const sheet =
      getActivitySheetById_(
        sheetId
      );


    lectureColumn =
      Number(lectureColumn);


    if (lectureColumn < 5) {

      throw new Error(
        "المحاضرة غير صحيحة."
      );
    }


    personName =
      cleanName_(
        personName
      );


    if (!personName) {

      throw new Error(
        "يجب إدخال الاسم."
      );
    }


    /****************************************
     * لو Row محدد
     ****************************************/

    if (rowNumber) {

      rowNumber =
        Number(rowNumber);


      if (rowNumber < 2) {

        throw new Error(
          "رقم الصف غير صحيح."
        );
      }


      const storedName =
        cleanName_(
          sheet
            .getRange(
              rowNumber,
              2
            )
            .getValue()
        );


      if (!storedName) {

        throw new Error(
          "الاسم غير موجود."
        );
      }


      // منع حضور مكرر
      const cell =
        sheet
          .getRange(
            rowNumber,
            lectureColumn
          );


      const oldValue =
        cell.getValue();


      if (
        oldValue !== "" &&
        oldValue !== null
      ) {

        return {

          success: false,

          alreadyAttended: true,

          errorCode:
            "ALREADY_ATTENDED",

          message:
            "تم إضافة " +
            storedName +
            " بالفعل",

          rowNumber:
            rowNumber,

          name:
            storedName,

          count:
            Number(
              sheet
                .getRange(
                  rowNumber,
                  3
                )
                .getValue()
            ) || 0,

          timestamp:
            normalizeTimestampValue_(oldValue)
        };
      }


      const timestamp =
        getAttendanceTimestamp_();


      cell.setValue(
        timestamp
      );


      const countCell =
        sheet
          .getRange(
            rowNumber,
            3
          );


      const oldCount =
        Number(
          countCell.getValue()
        ) || 0;


      const newCount =
        oldCount + 1;


      countCell.setValue(
        newCount
      );


      return {

        success: true,

        alreadyAttended: false,

        rowNumber:
          rowNumber,

        name:
          storedName,

        id:
          String(
            sheet
              .getRange(
                rowNumber,
                1
              )
              .getValue()
          ).trim(),

        count:
          newCount,

        timestamp:
          timestamp
      };
    }


    /****************************************
     * بدون Row
     ****************************************/

    const matches =
      findRowsByName_(
        sheet,
        personName
      );


    // أكثر من شخص بنفس الاسم
    if (matches.length > 1) {

      return {

        success: false,

        requiresSelection: true,

        matches:
          matches

      };
    }


    // شخص موجود
    if (matches.length === 1) {

      return recordManualAttendance(
        sheetId,
        lectureColumn,
        personName,
        matches[0].rowNumber
      );
    }


    /****************************************
     * شخص جديد بدون ID
     ****************************************/

    const newRow =
      createPersonRow_(
        sheet,
        "",
        personName
      );


    const timestamp =
      getAttendanceTimestamp_();


    sheet
      .getRange(
        newRow,
        lectureColumn
      )
      .setValue(
        timestamp
      );


    sheet
      .getRange(
        newRow,
        3
      )
      .setValue(
        1
      );


    return {

      success: true,

      alreadyAttended: false,

      rowNumber:
        newRow,

      id:
        "",

      name:
        personName,

      count:
        1,

      timestamp:
        timestamp,

      manual:
        true
    };


  } finally {

    lock.releaseLock();

  }
}


/**
 * البحث عن الاسم كامل
 */
function findRowsByName_(
  sheet,
  personName
) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        3
      )
      .getValues();


  const normalizedSearch =
    normalizeArabicName_(
      personName
    );


  const result = [];


  values.forEach(function (row, index) {

    const id =
      String(row[0] || "")
        .trim();


    const name =
      cleanName_(row[1]);


    if (
      normalizeArabicName_(name) ===
      normalizedSearch
    ) {

      result.push({

        rowNumber:
          index + 2,

        id:
          id,

        name:
          name,

        count:
          Number(row[2]) || 0,

        hasId:
          !!id

      });
    }

  });


  return result;
}


/*******************************************************
 * 17) ATTENDANCE LIST
 *******************************************************/

/**
 * كشف الحضور
 *
 * يعرض:
 * - أصحاب ID
 * - بدون ID
 * - أصحاب شهادة استثنائية حتى لو 0 حضور
 *
 * ملاحظة: يتم دائمًا إرجاع "cols" — خريطة { رقم عمود المحاضرة: التاريخ والوقت }
 * لكل الشخص، بغض النظر عن lectureColumn المُمرَّر، حتى تستطيع الواجهة
 * عرض تاريخ ووقت حضور كل شخص فى كل محاضرة (فى كشف الحضور الكامل)
 * أو فى محاضرة واحدة فقط (فى شاشة المحاضرة).
 */
function getAttendanceList(
  sheetId,
  lectureColumn
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();


  const result = [];


  values.forEach(function (row, index) {

    const id =
      String(row[0] || "")
        .trim();


    const name =
      cleanName_(row[1]);


    const count =
      Number(row[2]) || 0;


    const certificate =
      String(row[3] || "")
        .trim();


    if (!name) {
      return;
    }


    let attendedThisLecture =
      false;


    let timestamp = "";


    // خريطة تاريخ/وقت الحضور لكل محاضرة (كل الأعمدة من العمود 5 فصاعدًا)
    const cols = {};

    for (let c = 5; c <= lastColumn; c++) {

      const cellValue =
        row[c - 1];

      if (
        cellValue !== "" &&
        cellValue !== null
      ) {

        cols[c] =
          normalizeTimestampValue_(cellValue);
      }
    }


    if (
      lectureColumn &&
      Number(lectureColumn) >= 5 &&
      Number(lectureColumn) <= lastColumn
    ) {

      const value =
        row[
          Number(lectureColumn) - 1
        ];


      if (
        value !== "" &&
        value !== null
      ) {

        attendedThisLecture = true;

        timestamp =
          normalizeTimestampValue_(value);
      }
    }


    result.push({

      rowNumber:
        index + 2,

      id:
        id,

      name:
        name,

      count:
        count,

      certificate:
        certificate,

      hasId:
        !!id,

      attendedAny:
        count > 0,

      attendedThisLecture:
        attendedThisLecture,

      timestamp:
        timestamp,

      cols:
        cols
    });

  });


  result.sort(function (a, b) {

    return a.name.localeCompare(
      b.name,
      "ar"
    );

  });


  return result;
}


/*******************************************************
 * 18) FILTER ATTENDANCE
 *******************************************************/

/**
 * فلترة كشف الحضور
 *
 * filter:
 *
 * all
 * withId
 * withoutId
 */
function filterAttendanceList(
  sheetId,
  filter,
  search
) {

  let list =
    getAttendanceList(
      sheetId
    );


  filter =
    String(filter || "all");


  search =
    normalizeArabicName_(
      search
    );


  if (filter === "withId") {

    list =
      list.filter(function (item) {

        return item.hasId;

      });
  }


  if (filter === "withoutId") {

    list =
      list.filter(function (item) {

        return !item.hasId;

      });
  }


  if (search) {

    list =
      list.filter(function (item) {

        return normalizeArabicName_(
          item.name
        ).indexOf(search) !== -1;

      });
  }


  return list;
}


/*******************************************************
 * 19) CERTIFICATE NAME NORMALIZATION
 *******************************************************/

/**
 * Normalize اسم الشهادة
 */
function normalizeCertificateName_(
  name
) {

  let value =
    String(name || "")
      .trim();


  // إزالة .pdf
  value =
    value.replace(
      /\.pdf$/i,
      ""
    );


  // لو اسم الملف:
  // Name - CourseName
  //
  // نأخذ الجزء الأول
  const separators = [
    " - ",
    " – ",
    " — "
  ];


  for (let i = 0; i < separators.length; i++) {

    const separator =
      separators[i];


    const index =
      value.indexOf(
        separator
      );


    if (index > 0) {

      value =
        value.substring(
          0,
          index
        );

      break;
    }
  }


  return normalizeArabicName_(
    value
  );
}


/*******************************************************
 * 20) CERTIFICATE COURSE INFO
 *******************************************************/

/**
 * الحصول على بيانات الكورس للشهادة
 */
function getCourseInfoForCertificate_(
  courseId
) {

  return getCourseInfoById_(
    courseId
  );
}


/*******************************************************
 * 21) CHECK PDF URL
 *******************************************************/

/**
 * هل القيمة URL؟
 */
function isDriveUrl_(value) {

  if (!value) {
    return false;
  }


  const text =
    String(value).trim();


  return (
    text.indexOf("http://") === 0 ||
    text.indexOf("https://") === 0
  );
}


/*******************************************************
 * 21.1) CERTIFICATE FILE ID EXTRACTION + PDF FETCH
 *******************************************************/

/**
 * استخراج معرّف ملف Drive من رابط الشهادة
 *
 * يدعم الأشكال الشائعة:
 * https://drive.google.com/file/d/FILE_ID/view?usp=sharing
 * https://drive.google.com/open?id=FILE_ID
 * https://docs.google.com/.../d/FILE_ID/...
 */
function extractDriveFileId_(url) {

  url = String(url || "").trim();

  if (!url) {
    return null;
  }

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]{10,})/,
    /[?&]id=([a-zA-Z0-9_-]{10,})/
  ];

  for (let i = 0; i < patterns.length; i++) {

    const match = url.match(patterns[i]);

    if (match && match[1]) {
      return match[1];
    }
  }

  return null;
}


/**
 * جلب ملف الشهادة (PDF) كـ base64 حتى تستطيع الواجهة:
 * - عرضه داخل الموقع مباشرة (بدون التوجيه لرابط Drive)
 * - تنزيله كملف PDF فعلي
 * - مشاركته كملف PDF فعلي (وليس رابط)
 *
 * ملاحظة: اسم الملف اللي بيرجع هنا (fileName) هو مجرد fallback؛
 * الواجهة (Index.html) بتفرض اسم الشخص نفسه دايمًا كاسم للملف
 * وقت التنزيل/المشاركة، بدل الاعتماد على اسم الملف الداخلي فى Drive.
 */
function getCertificatePdf(url) {

  const fileId =
    extractDriveFileId_(url);

  if (!fileId) {

    throw new Error(
      "رابط الشهادة غير صالح."
    );
  }

  const file =
    DriveApp.getFileById(fileId);

  const blob =
    file.getBlob();

  const base64 =
    Utilities.base64Encode(
      blob.getBytes()
    );

  return {

    success: true,

    base64: base64,

    mimeType:
      blob.getContentType() ||
      "application/pdf",

    fileName:
      file.getName()

  };
}


/*******************************************************
 * 22) SYNC COURSE CERTIFICATES
 *******************************************************/

/**
 * مزامنة شهادات كورس
 *
 * Folder -> Sheet D
 */
function syncCourseCertificates(
  courseId
) {

  const course =
    getCourseInfoForCertificate_(
      courseId
    );


  if (!course) {

    throw new Error(
      "لم يتم العثور على الكورس."
    );
  }


  const sheetId =
    getCourseSheetId_(
      courseId
    );


  if (!sheetId) {

    throw new Error(
      "لم يتم العثور على Sheet الكورس."
    );
  }


  let sheet;


  try {

    sheet =
      getActivitySheetById_(
        sheetId
      );

  } catch (err) {

    throw new Error(
      "Sheet الكورس غير موجود."
    );
  }


  if (!course.folderId) {

    throw new Error(
      "لا يوجد Folder للشهادات."
    );
  }


  const folder =
    DriveApp.getFolderById(
      course.folderId
    );


  /****************************************
   * قراءة ملفات PDF
   ****************************************/

  const files =
    folder.getFiles();


  const pdfs = [];


  while (files.hasNext()) {

    const file =
      files.next();


    if (
      file.getMimeType() ===
      MimeType.PDF
    ) {

      pdfs.push({

        id:
          file.getId(),

        name:
          file.getName(),

        url:
          file.getUrl()

      });
    }
  }


  /****************************************
   * بناء Map بالأسماء
   ****************************************/

  const pdfMap = {};


  pdfs.forEach(function (pdf) {

    const normalized =
      normalizeCertificateName_(
        pdf.name
      );


    if (!normalized) {
      return;
    }


    if (!pdfMap[normalized]) {
      pdfMap[normalized] = [];
    }


    pdfMap[normalized].push(
      pdf
    );

  });


  /****************************************
   * قراءة Sheet
   ****************************************/

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return {

      success: true,

      courseId:
        courseId,

      matched:
        0,

      cleared:
        0,

      scannedFiles:
        pdfs.length,

      ambiguous: [],

      notFound: []

    };
  }


  const values =
    sheet
      .getRange(
        2,
        2,
        lastRow - 1,
        3
      )
      .getValues();


  let matched = 0;
  let cleared = 0;


  const ambiguous = [];
  const notFound = [];


  values.forEach(function (
    row,
    index
  ) {

    const name =
      cleanName_(
        row[0]
      );


    const certificate =
      String(
        row[2] || ""
      ).trim();


    if (!name) {
      return;
    }


    // "لا يوجد شهادة"
    // لا يتم تغييره
    if (
      certificate ===
      "لا يوجد شهادة"
    ) {
      return;
    }


    const normalized =
      normalizeCertificateName_(
        name
      );


    const matches =
      pdfMap[normalized] || [];


    /****************************************
     * أكثر من PDF لنفس الاسم
     ****************************************/

    if (matches.length > 1) {

      ambiguous.push({

        rowNumber:
          index + 2,

        name:
          name,

        files:
          matches.map(
            function (item) {
              return item.name;
            }
          )

      });

      return;
    }


    /****************************************
     * PDF موجود
     ****************************************/

    if (matches.length === 1) {

      const newUrl =
        matches[0].url;


      if (
        certificate !== newUrl
      ) {

        sheet
          .getRange(
            index + 2,
            4
          )
          .setValue(
            newUrl
          );
      }


      matched++;

      return;
    }


    /****************************************
     * PDF غير موجود
     *
     * لو كان D فيه رابط قديم
     * يتم مسحه
     ****************************************/

    if (
      isDriveUrl_(
        certificate
      )
    ) {

      sheet
        .getRange(
          index + 2,
          4
        )
        .clearContent();


      cleared++;

      return;
    }


    // قيمة يدوية أخرى
    if (
      certificate &&
      certificate !==
      "لا يوجد شهادة"
    ) {

      return;
    }


    notFound.push({

      rowNumber:
        index + 2,

      name:
        name

    });

  });


  return {

    success: true,

    courseId:
      courseId,

    courseName:
      course.name,

    matched:
      matched,

    cleared:
      cleared,

    scannedFiles:
      pdfs.length,

    ambiguous:
      ambiguous,

    notFound:
      notFound

  };
}


/*******************************************************
 * 23) SYNC ALL CERTIFICATES
 *******************************************************/

/**
 * مزامنة كل الكورسات
 */
function syncAllCourseCertificates() {

  const coursesSheet =
    getCoursesSheet_();


  const lastRow =
    coursesSheet.getLastRow();


  const summary = {

    success: true,

    courses: 0,

    matched: 0,

    cleared: 0,

    scannedFiles: 0,

    errors: []

  };


  if (lastRow < 2) {
    return summary;
  }


  const courseIds =
    coursesSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  courseIds.forEach(function (row) {

    const courseId =
      row[0];


    if (!courseId) {
      return;
    }


    summary.courses++;


    try {

      const result =
        syncCourseCertificates(
          courseId
        );


      summary.matched +=
        result.matched || 0;


      summary.cleared +=
        result.cleared || 0;


      summary.scannedFiles +=
        result.scannedFiles || 0;


    } catch (err) {

      summary.errors.push({

        courseId:
          courseId,

        error:
          err.message

      });
    }

  });


  return summary;
}


/*******************************************************
 * 24) TEST CERTIFICATE SYNC
 *******************************************************/

function testSyncCourseCertificates(
  courseId
) {

  return syncCourseCertificates(
    courseId
  );
}


/*******************************************************
 * 25) COURSE NAME SYNC
 *******************************************************/

/**
 * مزامنة أسماء الكورسات
 *
 * المصدر الأساسي:
 * Courses
 *
 * يتم تحديث:
 * - Attendance Sheet
 * - Drive Folder
 *
 * بدون حذف أو إنشاء جديد
 */
function syncCourseNamesFromCourses() {

  const coursesSheet =
    getCoursesSheet_();


  const lastRow =
    coursesSheet.getLastRow();


  if (lastRow < 2) {

    return {

      success: true,

      updated: 0

    };
  }


  const rows =
    coursesSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        4
      )
      .getValues();


  let updated = 0;


  rows.forEach(function (row) {

    const courseId =
      row[0];


    const courseName =
      cleanName_(row[1]);


    const folderId =
      String(row[2] || "")
        .trim();


    if (
      !courseId ||
      !courseName
    ) {
      return;
    }


    const sheetId =
      getCourseSheetId_(
        courseId
      );


    if (sheetId) {

      try {

        const sheet =
          getActivitySheetById_(
            sheetId
          );


        if (
          sheet.getName() !==
          courseName
        ) {

          sheet.setName(
            courseName
          );

          updated++;

        }

      } catch (err) {

        // Sheet اتحذفت يدويًا
        // لا ننشئ واحدة جديدة
      }
    }


    /****************************************
     * Rename Folder
     ****************************************/

    if (folderId) {

      try {

        const folder =
          DriveApp.getFolderById(
            folderId
          );


        if (
          folder.getName() !==
          courseName
        ) {

          folder.setName(
            courseName
          );

          updated++;

        }

      } catch (err) {

        // Folder غير موجود
      }
    }

  });


  return {

    success: true,

    updated:
      updated

  };
}


/*******************************************************
 * 26) COURSES EDIT TRIGGER
 *******************************************************/

/**
 * Trigger يتم تشغيله عند تعديل Courses
 */
function onCoursesEdit(e) {

  if (!e || !e.range) {
    return;
  }


  const sheet =
    e.range.getSheet();


  if (
    sheet.getName() !==
    SYSTEM_CONFIG.COURSES_SHEET
  ) {
    return;
  }


  // نعمل المزامنة كاملة
  // لأن تغيير اسم كورس أو إضافة كورس
  // يجب أن ينعكس على النظام
  try {

    syncCourseNamesFromCourses();

  } catch (err) {

    console.error(
      err.message
    );

  }
}


/*******************************************************
 * 27) SETUP COURSE EDIT TRIGGER
 *******************************************************/

/**
 * تشغيلها مرة واحدة فقط
 */
function setupCoursesEditTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();


  triggers.forEach(function (trigger) {

    if (
      trigger.getHandlerFunction() ===
      "onCoursesEdit"
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );
    }

  });


  ScriptApp
    .newTrigger(
      "onCoursesEdit"
    )
    .forSpreadsheet(
      SYSTEM_CONFIG.LEGACY_SPREADSHEET_ID
    )
    .onEdit()
    .create();


  return {

    success: true,

    message:
      "تم إنشاء Trigger لمزامنة أسماء الكورسات."

  };
}


/*******************************************************
 * 28) CERTIFICATE AUTO SYNC TRIGGER
 *******************************************************/

/**
 * تشغيلها مرة واحدة فقط
 *
 * مزامنة الشهادات كل دقيقة
 */
function setupCertificatesSyncTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();


  triggers.forEach(function (trigger) {

    if (
      trigger.getHandlerFunction() ===
      "syncAllCourseCertificates"
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );
    }

  });


  ScriptApp
    .newTrigger(
      "syncAllCourseCertificates"
    )
    .timeBased()
    .everyMinutes(1)
    .create();


  return {

    success: true,

    message:
      "تم إنشاء Trigger لمزامنة الشهادات كل دقيقة."

  };
}


/*******************************************************
 * 29) SETUP ALL TRIGGERS
 *******************************************************/

/**
 * شغل هذه الدالة مرة واحدة
 * بعد تركيب الكود
 */
function setupAllTriggers() {

  const courseTrigger =
    setupCoursesEditTrigger();


  const certificateTrigger =
    setupCertificatesSyncTrigger();


  return {

    success: true,

    courseTrigger:
      courseTrigger,

    certificateTrigger:
      certificateTrigger

  };
}


/*******************************************************
 * 30) OLD CERTIFICATE WEBSITE
 *******************************************************/

/**
 * doGet
 *
 * يحافظ على الموقع القديم
 */
function doGet() {

  return HtmlService
    .createTemplateFromFile(
      "Index"
    )
    .evaluate()
    .setTitle(
      "إصدار الشهادات"
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


/**
 * Include HTML
 */
function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(
      filename
    )
    .getContent();
}


/*******************************************************
 * 31) GET LEGACY CERTIFICATES
 *******************************************************/

/**
 * قراءة الشهادات القديمة
 */
function getLegacyCertificates_() {

  const sheet =
    getLegacyCertificatesSheet_();


  if (!sheet) {
    return [];
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        3
      )
      .getValues();


  const result = [];


  values.forEach(function (row) {

    const courseId =
      row[0];


    const name =
      cleanName_(
        row[1]
      );


    const pdf =
      String(
        row[2] || ""
      ).trim();


    if (
      !name ||
      !pdf
    ) {
      return;
    }


    const course =
      getCourseInfoById_(
        courseId
      );


    result.push({

      source:
        "legacy",

      courseId:
        courseId,

      courseName:
        course
          ? course.name
          : "",

      name:
        name,

      pdf:
        pdf

    });

  });


  return result;
}


/*******************************************************
 * 32) GET NEW CERTIFICATES
 *******************************************************/

/**
 * قراءة شهادات نظام الحضور الجديد
 */
function getNewCertificates_() {

  const coursesSheet =
    getCoursesSheet_();


  const lastRow =
    coursesSheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const rows =
    coursesSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        4
      )
      .getValues();


  const result = [];


  rows.forEach(function (row) {

    const courseId =
      row[0];


    const courseName =
      cleanName_(
        row[1]
      );


    if (
      !courseId ||
      !courseName
    ) {
      return;
    }


    const sheetId =
      getCourseSheetId_(
        courseId
      );


    if (!sheetId) {
      return;
    }


    let sheet;


    try {

      sheet =
        getActivitySheetById_(
          sheetId
        );

    } catch (err) {

      return;
    }


    const lastRow =
      sheet.getLastRow();


    if (lastRow < 2) {
      return;
    }


    const values =
      sheet
        .getRange(
          2,
          2,
          lastRow - 1,
          3
        )
        .getValues();


    values.forEach(function (item) {

      const name =
        cleanName_(
          item[0]
        );


      const certificate =
        String(
          item[2] || ""
        ).trim();


      if (
        !name ||
        !certificate
      ) {
        return;
      }


      if (
        certificate ===
        "لا يوجد شهادة"
      ) {
        return;
      }


      if (
        !isDriveUrl_(
          certificate
        )
      ) {
        return;
      }


      result.push({

        source:
          "attendance",

        courseId:
          courseId,

        courseName:
          courseName,

        name:
          name,

        pdf:
          certificate

      });

    });

  });


  return result;
}


/*******************************************************
 * 33) GET CERTIFICATES
 *******************************************************/

/**
 * الموقع القديم يستخدم هذه الدالة
 *
 * تجمع:
 *
 * 1) Certificates القديمة
 * 2) شهادات Sheets الجديدة
 */
function getCertificates() {

  const legacy =
    getLegacyCertificates_();


  const modern =
    getNewCertificates_();


  return legacy.concat(
    modern
  );
}


/*******************************************************
 * 34) TEST FUNCTIONS
 *******************************************************/

/**
 * اختبار الاتصال
 */
function testBackend() {

  const attendance =
    getAttendanceSpreadsheet_();


  const legacy =
    getLegacySpreadsheet_();


  return {

    success: true,

    attendanceSpreadsheet:
      attendance.getName(),

    legacySpreadsheet:
      legacy.getName(),

    coursesSheet:
      getCoursesSheet_().getName(),

    certificatesParentFolder:
      getCertificatesParentFolder_()
        .getName()

  };
}


/**
 * اختبار Folder الشهادات
 */
function testCertificatesFolder() {

  const folder =
    getCertificatesParentFolder_();


  return {

    success: true,

    id:
      folder.getId(),

    name:
      folder.getName(),

    url:
      folder.getUrl()

  };
}


/**
 * اختبار مزامنة أسماء الكورسات
 */
function testSyncCourseNames() {

  return syncCourseNamesFromCourses();
}


/**
 * اختبار كل الشهادات
 */
function testSyncAllCertificates() {

  return syncAllCourseCertificates();
}


/*******************************************************
 * 35) GET SINGLE ACTIVITY
 *******************************************************/

/**
 * الحصول على نشاط واحد
 */
function getActivity(sheetId) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  const courseId =
    getCourseIdBySheetId_(
      sheetId
    );


  const isMeeting =
    isMeetingSheet_(
      sheetId
    );


  const summary =
    getActivitySummary_(
      sheet
    );


  return {

    success: true,

    type:
      courseId
        ? "course"
        : isMeeting
          ? "meeting"
          : "activity",

    courseId:
      courseId,

    name:
      sheet.getName(),

    sheetId:
      sheet.getSheetId(),

    createdAt:
      getActivityCreatedAt_(
        sheetId
      ),

    createdDate:
      formatActivityDate_(
        getActivityCreatedAt_(
          sheetId
        )
      ),

    peopleCount:
      summary.peopleCount,

    actualAttendees:
      summary.actualAttendees,

    lectures:
      summary.lectures

  };
}


/*******************************************************
 * 36) GET LECTURE ATTENDANCE
 *******************************************************/

/**
 * كشف حضور محاضرة معينة
 */
function getLectureAttendance(
  sheetId,
  lectureColumn
) {

  const list =
    getAttendanceList(
      sheetId,
      lectureColumn
    );


  return list.filter(
    function (item) {

      return item.attendedThisLecture;

    }
  );
}


/*******************************************************
 * 37) ADD PERSON MANUALLY
 *******************************************************/

/**
 * إضافة شخص للـSheet بدون تسجيل حضور
 *
 * تستخدم في حالة:
 * - شهادة استثنائية
 * - شخص مطلوب يظهر في الكشف
 * - C = 0
 */
function addPersonWithoutAttendance(
  sheetId,
  personName,
  personId
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const sheet =
      getActivitySheetById_(
        sheetId
      );


    personName =
      cleanName_(
        personName
      );


    personId =
      String(
        personId || ""
      ).trim();


    if (!personName) {

      throw new Error(
        "يجب إدخال الاسم."
      );
    }


    // لو فيه ID
    if (personId) {

      const existingRow =
        findRowById_(
          sheet,
          personId
        );


      if (existingRow) {

        throw new Error(
          "هذا الـID موجود بالفعل."
        );
      }
    }


    // إضافة الاسم فقط
    const rowNumber =
      createPersonRow_(
        sheet,
        personId,
        personName
      );


    return {

      success: true,

      rowNumber:
        rowNumber,

      id:
        personId,

      name:
        personName,

      count:
        0,

      certificate:
        ""

    };


  } finally {

    lock.releaseLock();

  }
}


/*******************************************************
 * 38) UPDATE PERSON ID
 *******************************************************/

/**
 * إضافة ID لاحقًا لشخص كان بدون ID
 */
function updatePersonId(
  sheetId,
  rowNumber,
  personId
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const sheet =
      getActivitySheetById_(
        sheetId
      );


    rowNumber =
      Number(rowNumber);


    personId =
      String(
        personId || ""
      ).trim();


    if (
      rowNumber < 2 ||
      !personId
    ) {

      throw new Error(
        "بيانات غير صحيحة."
      );
    }


    const existingRow =
      findRowById_(
        sheet,
        personId
      );


    if (
      existingRow &&
      existingRow !== rowNumber
    ) {

      throw new Error(
        "هذا الـID مستخدم بالفعل."
      );
    }


    sheet
      .getRange(
        rowNumber,
        1
      )
      .setValue(
        personId
      );


    return {

      success: true,

      rowNumber:
        rowNumber,

      id:
        personId

    };


  } finally {

    lock.releaseLock();

  }
}


/*******************************************************
 * 39) UPDATE CERTIFICATE MANUALLY
 *******************************************************/

/**
 * تعديل قيمة الشهادة يدويًا
 *
 * القيمة ممكن تكون:
 * - PDF URL
 * - لا يوجد شهادة
 */
function updateCertificateValue(
  sheetId,
  rowNumber,
  value
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  rowNumber =
    Number(rowNumber);


  if (rowNumber < 2) {

    throw new Error(
      "رقم الصف غير صحيح."
    );
  }


  value =
    String(
      value || ""
    ).trim();


  sheet
    .getRange(
      rowNumber,
      4
    )
    .setValue(
      value
    );


  return {

    success: true,

    rowNumber:
      rowNumber,

    value:
      value

  };
}


/*******************************************************
 * 40) DELETE / CLEAR CERTIFICATE VALUE
 *******************************************************/

function clearCertificateValue(
  sheetId,
  rowNumber
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  rowNumber =
    Number(rowNumber);


  if (rowNumber < 2) {

    throw new Error(
      "رقم الصف غير صحيح."
    );
  }


  sheet
    .getRange(
      rowNumber,
      4
    )
    .clearContent();


  return {

    success: true

  };
}


/*******************************************************
 * 41) GET COURSE BY ID
 *******************************************************/

function getCourseById(
  courseId
) {

  const course =
    getCourseInfoById_(
      courseId
    );


  if (!course) {

    return null;
  }


  const sheetId =
    getCourseSheetId_(
      courseId
    );


  let sheet = null;


  if (sheetId) {

    try {

      sheet =
        getActivitySheetById_(
          sheetId
        );

    } catch (err) {

      sheet = null;

    }
  }


  return {

    courseId:
      course.courseId,

    name:
      course.name,

    folderId:
      course.folderId,

    folderUrl:
      course.folderUrl,

    sheetId:
      sheet
        ? sheet.getSheetId()
        : null,

    createdAt:
      sheet
        ? getActivityCreatedAt_(
            sheet.getSheetId()
          )
        : "",

    createdDate:
      sheet
        ? formatActivityDate_(
            getActivityCreatedAt_(
              sheet.getSheetId()
            )
          )
        : ""

  };
}


/*******************************************************
 * 42) GET ACTIVITY PEOPLE COUNT
 *******************************************************/

function getActivityPeopleCount(
  sheetId
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return 0;
  }


  const names =
    sheet
      .getRange(
        2,
        2,
        lastRow - 1,
        1
      )
      .getValues();


  return names.filter(
    function (row) {

      return cleanName_(
        row[0]
      ) !== "";

    }
  ).length;
}


/*******************************************************
 * 43) GET ACTUAL ATTENDANCE COUNT
 *******************************************************/

function getActualAttendanceCount(
  sheetId
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return 0;
  }


  const values =
    sheet
      .getRange(
        2,
        2,
        lastRow - 1,
        2
      )
      .getValues();


  return values.filter(
    function (row) {

      const name =
        cleanName_(row[0]);

      const count =
        Number(row[1]) || 0;


      return (
        name &&
        count > 0
      );

    }
  ).length;
}


/*******************************************************
 * 44) REMOVE EMPTY / INVALID DUPLICATE IDS
 *     NOT AUTOMATICALLY RUN
 *******************************************************/

/**
 * فحص الـIDs المكررة
 *
 * لا يقوم بالحذف.
 * فقط يعيد النتائج.
 */
function findDuplicateIds(
  sheetId
) {

  const sheet =
    getActivitySheetById_(
      sheetId
    );


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  const map = {};
  const duplicates = [];


  values.forEach(function (row, index) {

    const id =
      String(row[0] || "")
        .trim();


    const name =
      cleanName_(
        row[1]
      );


    if (!id) {
      return;
    }


    if (!map[id]) {

      map[id] = [];

    }


    map[id].push({

      rowNumber:
        index + 2,

      name:
        name

    });

  });


  Object.keys(map).forEach(
    function (id) {

      if (map[id].length > 1) {

        duplicates.push({

          id:
            id,

          records:
            map[id]

        });

      }

    }
  );


  return duplicates;
}


/*******************************************************
 * 45) INITIAL SETUP
 *******************************************************/

/**
 * شغلها بعد وضع الكود أول مرة
 *
 * لا تنشئ بيانات.
 * فقط:
 * - تزامن أسماء الكورسات
 * - تنشئ Triggers
 */
function initialSetup() {

  const syncNames =
    syncCourseNamesFromCourses();


  const triggers =
    setupAllTriggers();


  return {

    success: true,

    syncNames:
      syncNames,

    triggers:
      triggers

  };
}
/*******************************************************
 * 46) OFFLINE SUPPORT — لقطة للتخزين المحلي + مزامنة دفعة
 *
 * الفكرة:
 * - التطبيق (لما يكون أونلاين) بيسحب "لقطة" (snapshot) من نشاط معيّن
 *   ويخزّنها محليًا على الجهاز (روستر الأسماء + الـIDs + كل الحضور
 *   السابق + الشهادات).
 * - وهو أوفلاين: السكان والإضافة اليدوية ومنع التكرار كلها بتشتغل
 *   على النسخة المحلية دي فقط، وكل تسجيل حضور بيتخزن محليًا فى
 *   طابور "لسه متزامنش" (pending queue) بمعرف فريد (localId) من
 *   عند التطبيق نفسه.
 * - لما يرجع النت: التطبيق بيبعت كل الطابور دفعة واحدة لـ
 *   batchSyncAttendance، ويستقبل رد لكل عنصر فيه (نجح / مكرر /
 *   يحتاج اختيار يدوي / خطأ) عشان يحدّث حالته محليًا.
 *******************************************************/

/**
 * لقطة نشاط واحد صالحة للتخزين المحلي (أوفلاين).
 * نداء واحد بدل نداءين (getActivity + getAttendanceList) لتقليل
 * عدد رحلات الشبكة وقت المزامنة الأولى.
 */
function getActivitySnapshot(sheetId) {

  const activity =
    getActivity(sheetId);

  const people =
    getAttendanceList(
      sheetId,
      null
    );

  return {

    success: true,

    activity: activity,

    people: people,

    // وقت أخذ اللقطة (وقت الخادم) — يفيد التطبيق فى معرفة
    // مدى قِدَم النسخة المخزّنة محليًا
    snapshotAt:
      new Date().toISOString()

  };
}


/**
 * مزامنة دفعة من تسجيلات حضور اتعملت أوفلاين (سكان أو يدوي)
 * لمحاضرة واحدة، دفعة واحدة، بترتيب حدوثها الفعلي.
 *
 * lectureColumn: رقم عمود المحاضرة (كل عناصر الدفعة لنفس المحاضرة)
 *
 * records: مصفوفة عناصر:
 * {
 *   localId:         معرف فريد من عند التطبيق (مثلاً uuid) — إلزامي
 *   type:            "qr" أو "manual"
 *   qrData:          نص QR كامل — مطلوب لو type = "qr"
 *   personName:      اسم الشخص — مطلوب لو type = "manual" وrowNumber مش معروف
 *   rowNumber:       رقم الصف لو الشخص معروف بالفعل من اللقطة المحلية (اختياري)
 *   personId:        ID اختياري لو الشخص جديد
 *   clientTimestamp: وقت التسجيل الفعلي أوفلاين (ISO string) — لو فاضي بيتاخد وقت الآن
 * }
 *
 * الرد: مصفوفة نتائج بنفس عدد وترتيب records، كل عنصر:
 * {
 *   localId,
 *   status: "success" | "duplicate" | "ambiguous" | "error",
 *   name, rowNumber, count, timestamp,   // عند success/duplicate
 *   matches,                              // عند ambiguous — يحتاج المستخدم يختار يدويًا
 *   message                               // عند duplicate/ambiguous/error
 * }
 */
function batchSyncAttendance(
  sheetId,
  lectureColumn,
  records
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const sheet =
      getActivitySheetById_(
        sheetId
      );


    lectureColumn =
      Number(lectureColumn);


    if (lectureColumn < 5) {

      throw new Error(
        "المحاضرة غير صحيحة."
      );
    }


    if (!records || !records.length) {
      return [];
    }


    const results = [];


    records.forEach(function (record) {

      const localId =
        record.localId;

      try {

        const clientDate =
          record.clientTimestamp
            ? new Date(record.clientTimestamp)
            : new Date();

        const timestamp =
          formatAttendanceTimestamp_(
            clientDate
          );


        /*******************************
         * تسجيل عن طريق QR
         *******************************/

        if (record.type === "qr") {

          const person =
            parseQrData_(
              record.qrData
            );

          let rowNumber =
            findRowById_(
              sheet,
              person.id
            );


          if (rowNumber) {

            const currentName =
              cleanName_(
                sheet.getRange(rowNumber, 2).getValue()
              );

            const actualName =
              currentName || person.name;

            const cell =
              sheet.getRange(rowNumber, lectureColumn);

            const already =
              cell.getValue();


            if (already !== "" && already !== null) {

              results.push({
                localId: localId,
                status: "duplicate",
                name: actualName,
                rowNumber: rowNumber,
                count:
                  Number(sheet.getRange(rowNumber, 3).getValue()) || 0,
                timestamp: normalizeTimestampValue_(already),
                message:
                  "تم إضافة " + actualName + " بالفعل"
              });

              return;
            }


            if (!currentName) {
              sheet.getRange(rowNumber, 2).setValue(person.name);
            }


            cell.setValue(timestamp);

            const countCell =
              sheet.getRange(rowNumber, 3);

            const newCount =
              (Number(countCell.getValue()) || 0) + 1;

            countCell.setValue(newCount);


            results.push({
              localId: localId,
              status: "success",
              name: actualName,
              rowNumber: rowNumber,
              count: newCount,
              timestamp: timestamp
            });

            return;
          }


          // شخص جديد بالكامل (ID لم يُسجَّل من قبل)
          rowNumber =
            createPersonRow_(
              sheet,
              person.id,
              person.name
            );

          sheet.getRange(rowNumber, lectureColumn).setValue(timestamp);
          sheet.getRange(rowNumber, 3).setValue(1);


          results.push({
            localId: localId,
            status: "success",
            name: person.name,
            rowNumber: rowNumber,
            count: 1,
            timestamp: timestamp
          });

          return;
        }


        /*******************************
         * تسجيل يدوي
         *******************************/

        if (record.type === "manual") {

          let rowNumber =
            record.rowNumber
              ? Number(record.rowNumber)
              : null;

          const personName =
            cleanName_(record.personName);


          if (!personName && !rowNumber) {

            throw new Error(
              "بيانات التسجيل اليدوي غير مكتملة."
            );
          }


          // لو rowNumber مش معروف، حاول تلاقيه بالاسم
          if (!rowNumber) {

            const matches =
              findRowsByName_(sheet, personName);


            if (matches.length === 1) {

              rowNumber = matches[0].rowNumber;

            } else if (matches.length > 1) {

              // أكثر من شخص بنفس الاسم — لازم اختيار يدوي
              // بعد رجوع النت، مش تخمين من السيرفر
              results.push({
                localId: localId,
                status: "ambiguous",
                matches: matches,
                message:
                  "يوجد أكثر من شخص بهذا الاسم — يحتاج اختيار يدوي"
              });

              return;
            }
          }


          if (rowNumber) {

            const storedName =
              cleanName_(sheet.getRange(rowNumber, 2).getValue()) ||
              personName;

            const cell =
              sheet.getRange(rowNumber, lectureColumn);

            const already =
              cell.getValue();


            if (already !== "" && already !== null) {

              results.push({
                localId: localId,
                status: "duplicate",
                name: storedName,
                rowNumber: rowNumber,
                count:
                  Number(sheet.getRange(rowNumber, 3).getValue()) || 0,
                timestamp: normalizeTimestampValue_(already),
                message:
                  "تم إضافة " + storedName + " بالفعل"
              });

              return;
            }


            cell.setValue(timestamp);

            const countCell =
              sheet.getRange(rowNumber, 3);

            const newCount =
              (Number(countCell.getValue()) || 0) + 1;

            countCell.setValue(newCount);


            results.push({
              localId: localId,
              status: "success",
              name: storedName,
              rowNumber: rowNumber,
              count: newCount,
              timestamp: timestamp
            });

            return;
          }


          // شخص جديد بدون ID
          const newRow =
            createPersonRow_(
              sheet,
              record.personId || "",
              personName
            );

          sheet.getRange(newRow, lectureColumn).setValue(timestamp);
          sheet.getRange(newRow, 3).setValue(1);


          results.push({
            localId: localId,
            status: "success",
            name: personName,
            rowNumber: newRow,
            count: 1,
            timestamp: timestamp
          });

          return;
        }


        throw new Error(
          "نوع تسجيل غير معروف: " + record.type
        );


      } catch (err) {

        results.push({
          localId: localId,
          status: "error",
          message: err.message
        });
      }

    });


    return results;


  } finally {

    lock.releaseLock();

  }
}


/*******************************************************
 * 47) FIX EXISTING TIMESTAMP FORMATTING (تشغيل مرة واحدة)
 *
 * شغّلها مرة واحدة يدويًا من محرر Apps Script بعد رفع الكود لو
 * عندك بيانات حضور قديمة اتسجلت قبل التعديل ولسه شكلها غريب
 * (زي Wed Sep 23 2026 00:40:00 GMT+0300). الدالة دي بتظبط تنسيق
 * كل أعمدة المحاضرات فى كل الشيتات كـ "نص"، وتحوّل أي خلية لسه
 * متخزنة كـ Date حقيقي لنص منسّق بنفس شكل باقي التوقيتات.
 *******************************************************/
function fixExistingTimestampFormatting_() {

  const ss = getAttendanceSpreadsheet_();

  let fixedCells = 0;

  ss.getSheets().forEach(function (sheet) {

    // تظبيط التنسيق كـ نص لكل المستقبل
    sheet.getRange("E:ZZ").setNumberFormat("@");

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastColumn < 5 || lastRow < 2) {
      return;
    }

    const range =
      sheet.getRange(2, 5, lastRow - 1, lastColumn - 4);

    const values =
      range.getValues();

    let changed = false;

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {

        if (
          Object.prototype.toString.call(values[r][c]) ===
          "[object Date]"
        ) {

          values[r][c] =
            formatAttendanceTimestamp_(values[r][c]);

          changed = true;
          fixedCells++;
        }
      }
    }

    if (changed) {
      range.setValues(values);
    }

  });

  return {

    success: true,

    fixedCells: fixedCells

  };
}


function testBackendSimple() {

  try {

    const attendance =
      SpreadsheetApp.openById(
        SYSTEM_CONFIG.ATTENDANCE_SPREADSHEET_ID
      );

    const legacy =
      SpreadsheetApp.openById(
        SYSTEM_CONFIG.LEGACY_SPREADSHEET_ID
      );

    const courses =
      legacy.getSheetByName(
        SYSTEM_CONFIG.COURSES_SHEET
      );

    const folder =
      DriveApp.getFolderById(
        SYSTEM_CONFIG.CERTIFICATES_PARENT_FOLDER_ID
      );

    const result = {
      success: true,

      attendanceSpreadsheet:
        attendance.getName(),

      attendanceSpreadsheetId:
        attendance.getId(),

      legacySpreadsheet:
        legacy.getName(),

      legacySpreadsheetId:
        legacy.getId(),

      coursesSheet:
        courses
          ? courses.getName()
          : "NOT FOUND",

      certificatesFolder:
        folder.getName(),

      certificatesFolderId:
        folder.getId()
    };

    console.log(result);

    return result;

  } catch (error) {

    console.error(error);

    throw new Error(
      "Backend Test Error: " +
      error.message
    );
  }
}

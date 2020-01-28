// ==UserScript==
// @name         Canvas API Reports
// @namespace    https://github.com/djm60546/canvas-api-reports
// @version      1.1
// @description  Script for extracting student and instructor performance data using the Canvas API. Generates a .CSV download containing the data. Based on the Access Report Data script by James Jones.
// @author       Dan Murphy, Northwestern University School of Professional Studies (dmurphy@northwestern.edu)
// @match        https://canvas.northwestern.edu/accounts/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/1.3.3/FileSaver.js
// @require      https://code.jquery.com/jquery-3.4.1.js
// @require      https://code.jquery.com/ui/1.12.1/jquery-ui.js
// @grant        none

// ==/UserScript==

(function() {

    'use strict';

    // Some software doesn't like spaces in the headings.
    // Set headingNoSpaces = true to remove spaces from the headings
    var headingNoSpaces = false;

    // Viewing a student's profile page now counts as a participation.
    // This can confuse the faculty when student names show up as titles.
    // By default these are now removed from the data before downloading
    // Set showViewStudent = true to include these views in the data
    var showViewStudent = false;

    var userData = {};
    var currCourse = {};
    var enrollmentData = {};
    var assignmentData = {};
    var accessData = [];
    var atRiskArray = [];
    var instructorData = [];
    var submissionData = [];
    var tchrNameArray = [];
    var tchrEmailArray = [];
    var ajaxPool;
    var aborted = false;
    var topics = {};
    var topicEntries = [];
    var topicIDs = [];
    var topicsIdx = 0;
    var controls = {};;
    controls.accessCount = 0;
    controls.accessIndex = 0;
    controls.anonStdnts = true; // Anonymize student names and IDs
    controls.combinedRpt; // Make one report for all selected courses
    controls.canvasAcct = "21"; // SPS Canvas sub-account number
    controls.courseArray = [];
    controls.courseIndex = 0
    controls.coursePending = -1;
    controls.dlCrsOnly = true; // Make report(s) for online coures only
    controls.emptyCourse = false;
    controls.nowDate = new Date();
    controls.nowDateMS = controls.nowDate.getTime();
    controls.rptDateEnd = 0;
    controls.rptDateStart = 0;
    controls.rptDateStartTxt = '';
    controls.rptDateEndTxt = '';
    controls.rptType = 'none';
    controls.userArray = [];
    controls.userIndex = 0;
    controls.atRisk = {};
    // Criteria for students to be considered "at-risk'
    controls.atRisk.late = 0;
    controls.atRisk.mssg = 0
    controls.atRisk.time = 0;
    controls.atRisk.posts = 0;
    controls.atRisk.scoreRaw = 70.00; // students enrollment.current_score in Canvas
    controls.atRisk.sbmssn = 0;
    controls.lateGradingIntvl = 604800000 // 7 day period for on-time grades in milliseconds

    function errorHandler(e) {
        console.log(e.name + ': ' + e.message + 'at ' + e.stack);
        alert('An error occured. See browser console for details.');
        //wrapup();
    }

    function abortAll() {
        for (var i = 0; i < ajaxPool.length; i++) {
            ajaxPool[i].abort();
        }
        ajaxPool = [];
        wrapup();
    }

    function nextURL(linkTxt) {
        var url = null;
        if (linkTxt) {
            var links = linkTxt.split(',');
            var nextRegEx = new RegExp('^<(.*)>; rel="next"$');
            for (var i = 0; i < links.length; i++) {
                var matches = nextRegEx.exec(links[i]);
                if (matches) {
                    url = matches[1];
                }
            }
        }
        return url;
    }

    function setupPool() {
        try {
            ajaxPool = [];
            $.ajaxSetup({
                'beforeSend' : function(jqXHR) {
                    ajaxPool.push(jqXHR);
                },
                'complete' : function(jqXHR) {
                    var i = ajaxPool.indexOf(jqXHR);
                    if (i > -1) {
                        ajaxPool.splice(i, 1);
                    }
                }
            });
        } catch (e) {
            throw new Error('Error configuring AJAX pool');
        }
    }

    // Add current course data to a resource access or enrollment object
    function addCourseData(obj) {
        obj.sis_course_id = currCourse.sis_course_id;
        obj.course_code = currCourse.course_code;
        obj.course_id = currCourse.course_id;
        obj.course_name = currCourse.course_name;
        obj.enrollment_term_id = currCourse.enrollment_term_id;
        obj.quarter_name = currCourse.quarter_name[0];
        obj.section = currCourse.section[0];
        obj.short_course_code = currCourse.short_code[0];
        if (controls.rptType == 'at-risk' || controls.rptType == 'participation') {
            obj.teacher_name = currCourse.teacher_name;
            obj.teacher_email = currCourse.teacher_email;
        }
        if (controls.rptType == 'instructor') {
            obj.ttl_stdnts = currCourse.ttl_stdnts
            obj.graded_ontime_pcnt = currCourse.graded_ontime_pcnt;
            obj.graded_late_pcnt = currCourse.graded_late_pcnt;
            obj.graded_none_pcnt = currCourse.graded_none_pcnt;
            obj.feedback_count = currCourse.feedback_count;
            obj.feedback_mean_length = currCourse.feedback_mean_length;
        }
    }

    function getEnrollmentID(userID) {
        var enrollementID;
        for (var id in enrollmentData) {
            var enrollment = enrollmentData[id];
            if (enrollment.user_id == userID && enrollment.course_id == currCourse.course_id) {
                enrollementID = enrollmentData[id].id;
                break;
            }
        }
        return enrollementID;
    }

    function controller(state) {
        var sURL, aURL;
        controls.userIndex++;
        switch (state) {
            case 'accesses':
                if (controls.userIndex >= controls.userArray.length) {
                    processAccesses();
                } else {
                    aURL = '/courses/' + currCourse.course_id + '/users/' + controls.userArray[controls.userIndex] + '/usage.json?per_page=100';
                    getAccesses(aURL);
                }
                break;
            case 'course_done':
                progressbar(10, 10);
                controls.coursePending--;
                if (controls.combinedRpt && (controls.coursePending > 0)) { // if a combined report AND 1 or more courses pending
                    makeNewReport();
                } else { // if (a single report AND NO more courses pending) OR invidual reports
                    outputReport();
                }
                break;
            case 'enrollments_done':
                if (controls.rptType == 'access' || controls.rptType == 'at-risk' || controls.rptType == 'instructor') {
                    progressbar(0,controls.accessCount);
                    aURL = '/courses/' + currCourse.course_id + '/users/' + controls.userArray[controls.userIndex] + '/usage.json?per_page=100';
                    getAccesses(aURL);
                } else {
                    sURL = 'https://canvas.northwestern.edu/api/v1/courses/' + currCourse.course_id + '/students/submissions?student_ids[]=all&per_page=100';
                    getSubmissions(sURL);
                }
                break;
            case 'submissions':
                if (submissionData.length > 0) {
                    processStudentSubmissions();
                } else {
                    sURL = 'https://canvas.northwestern.edu/api/v1/courses/' + currCourse.course_id + '/students/submissions?student_ids[]=all&per_page=100';
                    sURL += controls.rptType == 'instructor' ? '&include[]=submission_comments' : '';
                    getSubmissions(sURL);
                }
                break;
            case 'topics':
                var tURL = '/api/v1/courses/' + currCourse.course_id + '/discussion_topics?per_page=100';
                getTopics(tURL);
                break;
            case 'instructor':
                for (var id in enrollmentData) {
                    var thisEnrollment = enrollmentData[id];
                    addCourseData(thisEnrollment);
                    instructorData.push(thisEnrollment);
                    //delete enrollmentData[user_id];
                }
                controller('course_done');
                break;

        }
    }

    function processTopicEntries() {
         $('#dm_report_status').text('Processing discussion topic entries...');
        for (var i = 0; i < topicEntries.length; i++) {
            progressbar(i, topicEntries.length);
            var thisEntry = topicEntries[i];
            var userID = thisEntry.user_id
            var thisEnrollment = enrollmentData[getEnrollmentID(userID)];
            if (typeof(thisEnrollment) != 'undefined') { // omit entries from missing or non-teacher enrollments;  && thisEnrollment
                var discussionCount = thisEnrollment.discussion_posts + 1;
                thisEnrollment.discussion_posts = discussionCount;
                var postChars = thisEnrollment.post_chars + thisEntry.message.length;
                thisEnrollment.post_chars = postChars;
                var meanPostLength = postChars / discussionCount;
                thisEnrollment.post_mean_length = meanPostLength;
                var thisUpdate = new Date(thisEntry.updated_at);
                var thisUpdateMS = thisUpdate.getTime();
                var lastPost = new Date(thisEnrollment.last_post_date);
                var lastPostMS = lastPost.getTime();
                if (thisUpdateMS > lastPostMS) {thisEnrollment.last_post_date = thisEntry.updated_at} // get most recent resource access
            }
        }
        controller('instructor');
    }

     // Get the list of discussion topics for the current couurse
    function getTopicEntries(eURL) {
         $('#dm_report_status').text('Getting discussion topic entries...');
        try {
            if (aborted) {
                throw new Error('Aborted at getTopicEntries');
            }
            $.getJSON(eURL, function(eData, status, jqXHR) {
                if (eData) {
                    for (var i = 0; i < eData.view.length; i++) {
                        progressbar(i, eData.view.length);
                        var thisEntry = eData.view[i];
                        if (thisEntry.hasOwnProperty('deleted') || !thisEntry.hasOwnProperty('message')) {continue} // omit deleted entries
                        // If reporting period indicated, get only those entries posted within the reporting period
                        var entryDate = new Date(thisEntry.updated_at);
                        var entryDateMS = entryDate.getTime();
                        if (controls.rptDateStart > 0 && (entryDateMS >= controls.rptDateStart && entryDateMS <= controls.rptDateEnd)) {continue}
                        topicEntries.push(thisEntry);
                          if (thisEntry.replies) {
                            var entryReplies = thisEntry.replies;
                            for (var j = 0; j < entryReplies.length; j++) {
                                var thisReply = entryReplies[j];
                                if (thisReply.hasOwnProperty('deleted') || !thisReply.hasOwnProperty('message')) { // omit deleted entries
                                    continue;
                                }
                                topicEntries.push(thisReply);
                            }
                        }
                    }
                }
            }).done(function () {
                getTopicEntryController();
            }).fail(function() {
                var errorDetail = 'Topic entries query for course ' + currCourse.course_id + ' threw an error';
                throw new Error(errorDetail);
                //controller(courseId);
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    function getTopicEntryController() {
        if (topicsIdx >= topicIDs.length) {
            processTopicEntries();
        } else {
            progressbar(topicsIdx, topicIDs.length);
            var currTopicID = topicIDs[topicsIdx];
            var eURL = '/api/v1/courses/' + currCourse.course_id + '/discussion_topics/' + currTopicID + '/view?per_page=100';
            getTopicEntries(eURL,currTopicID);
            topicsIdx++;
        }
    }

    // Get the list of discussion topics for the current couurse

    function getTopics(tURL) {
        try {
            if (aborted) {
                throw new Error('Aborted');
            }
            $('#dm_report_status').text('Getting discussion topics...');
            $.getJSON(tURL, function(tData, status, jqXHR) {
                tURL = nextURL(jqXHR.getResponseHeader('Link')); // Get next page of results, if any
                if (tData) {
                    for (var i = 0; i < tData.length; i++) {
                        progressbar(i, tData.length);
                        var thisTopic = tData[i];
                        topics[thisTopic.id] = thisTopic;
                        topicIDs.push(thisTopic.id);
                    }
                }
            }).done(function () {
                if (tURL) {
                    getTopics(tURL)
                } else {
                getTopicEntryController();
                }
            }).fail(function() {
                var errorDetail = 'Topics query for course ' + currCourse.course_id + ' threw an error';
                throw new Error(errorDetail);
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    function processGrading() {
        $('#dm_report_status').text('Processing submissions data...');
        var feedbackLen = 0;
        for (var i = 0; i < submissionData.length; i++) {
            progressbar(i, submissionData.length);
            var thisSubmission = submissionData[i];
            var thisAssignment = assignmentData[thisSubmission.assignment_id];
            if ((typeof(thisSubmission) != 'undefined' && thisSubmission) && (typeof(thisAssignment) != 'undefined' && thisAssignment)) {
                var dueDate = new Date(thisAssignment.due_at);
                var dueDateMS = dueDate.getTime();
                var submitDate = new Date(thisSubmission.submitted_at);
                var submitDateMS = submitDate.getTime();
                var gradedDate = new Date(thisSubmission.graded_at);
                var gradedDateMS = gradedDate.getTime();
                if ((dueDateMS > controls.nowDateMS) || thisSubmission.workflow_state == 'unsubmitted') {continue} // Exclude submissions due in the future or those not submitted
                currCourse.submissions_count++;
                if (thisSubmission.workflow_state == 'graded') {
                    if ((thisSubmission.late == false && gradedDateMS - dueDateMS <= controls.lateGradingIntvl) || (thisSubmission.late == true && gradedDateMS - submitDateMS <= controls.lateGradingIntvl)) { // Graded on time w/in 7 days after submission
                        currCourse.submissions_graded_ontime++
                    } else if ((thisSubmission.late == false && gradedDateMS - dueDateMS > controls.lateGradingIntvl) || (thisSubmission.late == true && gradedDateMS - submitDateMS > controls.lateGradingIntvl)) { // Graded late more than 7 days after submission
                        currCourse.submissions_graded_late++;
                    }
                } else if (thisSubmission.workflow_state == 'submitted' && ((controls.nowDateMS - dueDateMS > controls.lateGradingIntvl && thisSubmission.late == false) || (controls.nowDateMS - submitDateMS > controls.lateGradingIntvl && thisSubmission.late == true))) { // Not graded w/in 7 days after submission
                    currCourse.submissions_graded_none++;
                }
                for (var j = 0; j < thisSubmission.submission_comments.length; j++) {
                    var thisAuthor = userData[thisSubmission.submission_comments[j].author_id];
                    if (typeof(thisAuthor) == 'undefined' || !thisAuthor) {continue} // Skip comments from students
                    //console.log(thisSubmission.submission_comments[j]);
                    currCourse.feedback_count++;
                    var feedback = thisSubmission.submission_comments[j].comment;
                    feedbackLen += feedback.length;
                }
            }
        }
        var gradedNonePcnt = Math.round((currCourse.submissions_graded_none / currCourse.submissions_count) * 100, 2);
        currCourse.graded_none_pcnt = gradedNonePcnt;
        var gradedLatePcnt = Math.round((currCourse.submissions_graded_late / currCourse.submissions_count) * 100, 2);
        currCourse.graded_late_pcnt = gradedLatePcnt;
        var gradedOntimePcnt = Math.round((currCourse.submissions_graded_ontime / currCourse.submissions_count) * 100, 2);
        currCourse.graded_ontime_pcnt = gradedOntimePcnt;
        currCourse.feedback_mean_length = currCourse.feedback_count > 0 ? Math.round(feedbackLen / currCourse.feedback_count, 0) : 0;
        controller('topics');
    }

    // Make At-risk list based on the status and quantity students' sumbmissions, discussion posts and current score
    function processAtRisk() {
        $('#dm_report_status').text('Identifying at-risk students...');
        progressbar(9, 10);
        for (var id in enrollmentData) {
            var thisEnrollment = enrollmentData[id];
            if (thisEnrollment.role != 'Student') {continue}
            var submissions = thisEnrollment.submitted;
            var time = thisEnrollment.total_activity_time;
            var posts = thisEnrollment.discussion_posts;
            var late = thisEnrollment.assignments_late;
            var missing = thisEnrollment.assignments_missing;
            var score = thisEnrollment.grades.current_score;
            if (controls.rptType == 'at-risk') {
                if (late > controls.atRisk.late || missing > controls.atRisk.mssg || time === controls.atRisk.time || posts === controls.atRisk.posts || score == controls.atRisk.scoreRaw || submissions === controls.atRisk.sbmssn) {
                    thisEnrollment.current_score = score < controls.atRisk.scoreRaw ? 'Low' : 'OK';
                    thisEnrollment.discussion_posts = ' ' + posts + ' / ' + currCourse.discussions_due;
                    thisEnrollment.submitted = ' ' + submissions + ' / ' + currCourse.assignments_due;
                    addCourseData(thisEnrollment); // Add current course data to only those enrollments that will be reported
                    atRiskArray.push(thisEnrollment);
                }
            } else {
                if (submissions === 0) { // Zero participation criteria
                    thisEnrollment.submitted = ' ' + submissions + ' / ' + currCourse.assignments_due;
                    addCourseData(thisEnrollment); // Add current course data to only those enrollments that will be reported
                    atRiskArray.push(thisEnrollment);
                }

            }
        }
        controller('course_done');
    }

    function processStudentSubmissions() {
        $('#dm_report_status').text('Processing submissions data...');
        progressbar(8, 10);
        for (var i = 0; i < submissionData.length; i++) {
            var thisSubmission = submissionData[i];
            var userID = thisSubmission.user_id;
            var thisEnrollment = enrollmentData[getEnrollmentID(userID)];
            var thisAssignment = assignmentData[thisSubmission.assignment_id];
            if ((typeof(thisEnrollment) != 'undefined' && thisEnrollment) && (typeof(thisAssignment) != 'undefined' && thisAssignment)) {
                if (thisSubmission.workflow_state != 'unsubmitted'){
                    var submitted = thisEnrollment.submitted + 1;
                    thisEnrollment.submitted = submitted;
                }
                if (thisSubmission.submission_type == 'discussion_topic' && thisSubmission.workflow_state != 'unsubmitted'){
                    var discussions = thisEnrollment.discussion_posts + 1;
                    thisEnrollment.discussion_posts = discussions;
                }
                if (thisSubmission.late == true) {
                    var timeLate = thisSubmission.seconds_late;
                    if (timeLate > thisEnrollment.max_days_late) {thisEnrollment.max_days_late = timeLate}
                    var late = thisEnrollment.assignments_late + 1;
                    thisEnrollment.assignments_late = late;
                }
                if (thisSubmission.missing == true && thisSubmission.workflow_state == 'unsubmitted') {
                    var missing = thisEnrollment.assignments_missing + 1;
                    var timeMissing = thisSubmission.seconds_late;
                    if (timeMissing > thisEnrollment.max_days_missing) {thisEnrollment.max_days_missing = timeMissing}
                    thisEnrollment.assignments_missing = missing;
                }
            }
        }
        processAtRisk();
    }

    function getAssignments(dURL) {
        $('#dm_report_status').text('Getting assignments data...');
        progressbar(7, 10);
        // console.log('getAssignments');
        var discussions = 0;
        var assignments = 0;
        try {
            $('#dm_report_status').text('Getting assignments data...');
            $.getJSON(dURL, function(ddata, status, jqXHR) { // Get assignments for the current course
                dURL = nextURL(jqXHR.getResponseHeader('Link')); // Get next page or results, if any
                for (var i = 0; i < ddata.length; i++) {
                    var thisAssignment = ddata[i];
                    var due = new Date(thisAssignment.due_at);
                    var dueMS = due.getTime();
                    if (dueMS == 0) {continue}
                    // Get and count only those assignments that were due before today (if no reporting period specified) or those due within the reporting period
                    if ((controls.rptDateStart === 0 && dueMS < controls.nowDateMS) || (controls.rptDateStart > 0 && (dueMS >= controls.rptDateStart && dueMS <= controls.rptDateEnd))) {
                        assignmentData[thisAssignment.id] = thisAssignment;
                        assignments++;
                        if (thisAssignment.submission_types[0] == 'discussion_topic') {
                            discussions++;
                        }
                    }
                }
            }).done(function () {
                if (dURL) {
                    getAssignments(dURL);
                } else {
                    currCourse.assignments_due = assignments;
                    currCourse.discussions_due = discussions;
                    if (controls.rptType == 'instructor') {
                        processGrading();
                    } else {
                        processStudentSubmissions();
                    }
                }
            }).fail(function() {
                var errorDetail = 'Error getting assignments for course:' + currCourse.course_id;
                throw new Error(errorDetail);
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    function getSubmissions(sURL) {
        progressbar(5, 10);
        $('#dm_report_status').text('Getting submissions data...');
        try {
            $.getJSON(sURL, function(sdata, status, jqXHR) {
                sURL = nextURL(jqXHR.getResponseHeader('Link')); // Generate next page URL if more than 100 access records returned
                for (var i = 0; i < sdata.length; i++) {
                    submissionData.push(sdata[i]);
                }
            }).done(function() {
                if (sURL) {
                    getSubmissions(sURL);
                } else {
                    var dURL = '/api/v1/courses/' + currCourse.course_id + '/assignments?per_page=100';
                    getAssignments(dURL);
                }
            }).fail(function() {
                if (!aborted) {
                    var errorDetail = 'submissions,' + currCourse.id;
                    throw new Error(errorDetail);
                }
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    function writeNoAccesses(obj) {
        var nonAccess = {
            'role' : obj.role,
            'readable_name' : 'No accesses',
            'total_activity_time' : 0,
            'view_score' : 0,
            'participate_score' : 0,
            'created_at' : null,
            'last_activity_at' : null,
            'action_level' : 'none',
            'asset_category' : 'N/A',
            'asset_class_name' : 'N/A',
            'user_id' : obj.user_id
        };
        accessData.push(nonAccess);
    }

    function processAccesses() {
        $('#dm_report_status').text('Processing user access data...');
        for (var i = controls.accessIndex; i < accessData.length; i++) {
            var thisAccess = accessData[i];
            var userID = thisAccess.user_id;
            var thisEnrollment = enrollmentData[getEnrollmentID(userID)];
            if (typeof thisEnrollment !== 'undefined' && thisEnrollment) {
                if (controls.rptType == 'access') {
                    thisAccess.user_role = thisEnrollment.role;
                    thisAccess.total_activity_time = thisEnrollment.total_activity_time;
                } else if (controls.rptType == 'at-risk' || controls.rptType == 'instructor') {
                    var regex = RegExp('(.jpg|.png|.svg|.mp(3|4|g))$'); // filter out media files from count of resource views
                    var result = regex.test(thisAccess.readable_name);
                    if (!result) {thisEnrollment.page_views += thisAccess.view_score} // not media, so add to resource views
                    if (thisAccess.readable_name == 'Course Home') {thisEnrollment.home_page_views = thisAccess.view_score}
                    if (controls.rptDateStart > 0) {
                        var thisLastAccess = new Date(thisAccess.last_access);
                        var thisLastAccessMS = thisLastAccess.getTime();
                        var currLastAccess = new Date(thisEnrollment.last_access);;
                        var currLastAccessMS = currLastAccess.getTime()
                        if (thisLastAccessMS > currLastAccessMS) {thisEnrollment.last_activity_at = thisAccess.last_access} // get most recent resource access
                    }
               }
            } else {
                thisAccess.user_role = 'Unenrolled';
            }
            accessData[i] = thisAccess;
            addCourseData(accessData[i]);
        }
        controls.accessIndex = i;
        if (controls.rptType == 'access') {
            controller('course_done');
        } else {controller('submissions')}
    }

    function getAccesses(aURL) {
        $('#dm_report_status').text('Getting user access data...');
        try {
            $.getJSON(aURL, function(adata, status, jqXHR) {
                aURL = nextURL(jqXHR.getResponseHeader('Link')); // Generate next page URL if more than 100 access records returned
                if (adata.length === 0) {
                    var userID = controls.userArray[controls.userIndex];
                    var obj = enrollmentData[getEnrollmentID(userID)];
                    writeNoAccesses(obj);
                }
                for (var i = 0; i < adata.length; i++) {
                    var thisAccess = adata[i].asset_user_access;
                    var firstAccess = new Date(thisAccess.created_at);
                    var firstAccessMS = firstAccess.getTime();
                    var lastAccess = new Date(thisAccess.last_access);
                    var lastAccessMS = lastAccess.getTime();
                    if (controls.rptDateStart > 0 && (firstAccessMS >= controls.rptDateStart && lastAccessMS <= controls.rptDateEnd)) {continue}
                    accessData.push(thisAccess);
                }
            }).done(function() {
                if (aURL) {
                    getAccesses(aURL);
                } else {
                    progressbar(controls.userIndex, controls.userArray.length);
                    controller('accesses');
                }
            }).fail(function() {
                controller('accesses');
                if (!aborted) {
                    var errorDetail = 'access,' + controls.userArray[controls.userIndex] + ',' + currCourse.course_id;
                    throw new Error(errorDetail);
                }
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    // This function will anonymize the user list to protect student privacy
    // By randomizing first and last names seperately, a large number of student names can be randomly generated with minimal chance of duplication
    // First names are colors and last names are fruits, vegetables, spices, etc
    function anonStdnts() {
    var firstNames =
        ['Alabaster','Almond','Arylide','Ash','Beige','Bistre','Black','Bleu','Blizzard','Blood','Blue','Brass','Brick','Bronze','Brown','Cadmium','Cafe Au Lait','Canary','Carmine','Carnation','Cedar','Cerise','Cerulean','Champagne','Chartreuse','Chocolate','Chrome','Cinnamon','Cobalt','Cocoa','Coffee','Copper','Coral','Cornflower','Cotton','Cyan','Denim','Ecru','Emerald','Fallow','Falu','Fern','Flax','Forest','Fuchsia','Fulvous','Goldenrod','Gray','Green','Indigo','Khaki','Latte','Lava','Lavender','Lilac','Magenta','Maroon','Mauve','Olive','Opal','Orange','Pink','Purple','Raspberry','Red','Rose','Ruby','Saffron','Salmon','Sand','Sapphire','Satin','Sienna','Tangerine','Taupe','Turquoise','Umber','Vermillion','Violet','White','Yellow'
        ];
    var lastNames =
        ['Acorn','Alfalfa','Amrud','Anise','Artichoke','Arugula','Asparagus','Aubergine','Avacado','Azuki','Banana','Basil','Bean','Beet','Bok Choy','Borlotti','Broccoli','Cabbage','Caraway','Carrot','Cauliflower','Celeriac','Celery','Chamomile','Chard','Chestnut','Chickpea','Chive','Cilantro','Collard Green','Coriander','Cucumber','Daikon','Delicata','Dill','Eggplant','Endive','Fennel','Frisee','Garbanzo','Garlic','Ginger','Habanero','Horseradish','Jalapeño','Jicama','Kale','Kohlrabi','Lavender','Leek','Lemon','Lemon Grass','Lentils','Lettuce','Lima Bean','Mangel-Wurzel','Mangetout','Marjoram','Melon','Mushroom','Mustard','Nettles','Okra','Onion','Oregano','Paprika','Parsley','Parsnip','Pea','Pepper','Potato','Pumpkin','Quandong','Quinoa','Radicchio','Radish','Rhubarb','Rosemary','Rutabaga','Sage','Salsify','Scallion','Shallot','Skirret','Soy','Spinach','Sprout','Squash','Sunchoke','Sugar','Sweetcorn','Taro','Tat','Tat Soi','Thyme','Tomato','Topinambur','Turnip','Wasabi'
        ];
    var firstNamesLen = firstNames.length;
    var lastNamesLen = lastNames.length;
        for (var id in enrollmentData) {
            if (enrollmentData.hasOwnProperty(id)) {
                var userID = enrollmentData[id].user_id;
                // Do not anonymize the names of instructors, TAs or support. Do not give a new anonymized name to a student who has one.
                if (enrollmentData[id].role !== 'Student' || ((typeof userData[userID] == 'undefined' || userData[userID].anon_name))) {
                    continue;
                }
                var names = [];
                var firstNameIdx = Math.floor(Math.random() * firstNamesLen);
                var lastNameIdx = Math.floor(Math.random() * lastNamesLen);
                names.push(firstNames[firstNameIdx]);
                names.push(lastNames[lastNameIdx]);
                var name = names.join(' ');
                userData[userID].anon_name = name;
            }
        }
    }

    function processEnrollments() {
        $('#dm_report_status').text('Processing user access data...');
        progressbar(4, 10);
        var thisUserRole, nextFunction;
        for (var id in enrollmentData) {
            var thisEnrollment = enrollmentData[id];
            thisUserRole = thisEnrollment.role;
            thisUserRole = thisUserRole.replace(/Enrollment/i,""); // Delete 'Enrollment' from role
            thisUserRole = thisUserRole.replace(/^Ta$/,'TA'); // Capitalize both letters in TA user role
            thisEnrollment.role = thisUserRole;
            thisEnrollment.page_views = 0;
            thisEnrollment.home_page_views = 0;
            if (thisUserRole == "Student") {
                thisEnrollment.submitted = 0;
                thisEnrollment.assignments_late = 0;
                thisEnrollment.max_days_late = 0;
                thisEnrollment.assignments_missing = 0;
                thisEnrollment.max_days_missing = 0;
                thisEnrollment.discussion_posts = 0;
            } else if (thisUserRole == "Teacher" && (controls.rptType == 'at-risk' || controls.rptType == 'participation')) {
                var tchrID = thisEnrollment.user_id;
                var tchrName = userData[tchrID].name;
                if (!/^NU-/.test(tchrName)) { // Do not include admins enrolled as teachers using their "NU-[name]" accounts
                    var tchrEmail = userData[tchrID].email;
                    tchrEmailArray.push(tchrEmail);
                    tchrNameArray.push(tchrName);
                }
            } else if (controls.rptType == 'instructor') {
                thisEnrollment.discussion_posts = 0;
                thisEnrollment.post_chars = 0;
                thisEnrollment.post_mean_length = 0;
                thisEnrollment.last_post_date = 0;
            }
            enrollmentData[id] = thisEnrollment;
        }
        currCourse.teacher_email = tchrEmailArray.toString().replace(/,/g,';');
        currCourse.teacher_name = tchrNameArray.toString().replace(/,/g,', ');
        tchrEmailArray = [];
        tchrNameArray = [];
        if (controls.anonStdnts) {
            anonStdnts();
        }
        controller('enrollments_done');
    }

    // Get enrollment data for users that is course-specific
    function getEnrollments() {
        $('#dm_report_status').text('Getting enrollments data...');
        progressbar(3, 10);
        var eURL = '/api/v1/courses/' + currCourse.course_id + '/enrollments?per_page=100';
        eURL += controls.rptType == 'instructor' ? '&role[]=TeacherEnrollment' : '';
        try {
            if (aborted) {
                throw new Error('Aborted');
            }
            $.getJSON(eURL, function(edata, status, jqXHR) {
                if (edata) {
                    for (var i = 0; i < edata.length; i++) {
                        var thisEnrollment = edata[i];
                        enrollmentData[edata[i].id] = thisEnrollment;
                    }
                }
            }).done(function () {
                processEnrollments();
            }).fail(function() {
                var errorDetail = 'enrollment,' + currCourse.course_id;
                throw new Error(errorDetail);
                controller();
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    // Get the user data that is not course-specific
    function getUsers(uURL) {
        $('#dm_report_status').text('Getting users data...');
        progressbar(2, 10);
        try {
            $.getJSON(uURL, function(udata, status, jqXHR) { // Get users for the current course
                uURL = nextURL(jqXHR.getResponseHeader('Link')); // Get next page of results, if any
                for (var i = 0; i < udata.length; i++) {
                    var thisUser = udata[i];
                    userData[thisUser.id] = thisUser;
                    controls.userArray.push(thisUser.id);
                }
            }).fail(function() {
                var errorDetail = 'users,' + currCourse.course_id;
                throw new Error(errorDetail);
            }).done(function () {
                if (uURL) {
                    getUsers(uURL);
                } else {
                    getEnrollments();
                }
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    // Get course data one time to be applied to all users of the current course
    function getCourseData(crsID) {
        $('#dm_report_status').text('Getting course data...');
        progressbar(1,10); // Reset progress bar
        try {
            if (aborted) {
                throw new Error('Aborted at getCourseData: Course Info');
            }
            $('#dm_report_status').text('Getting course data...');
            var urlCrs = '/api/v1/courses/' + crsID + '?include[]=total_students';
            $.getJSON(urlCrs, function(cdata, status, jqXHR) {
                if (cdata) {
                    currCourse.sis_course_id = cdata.sis_course_id;
                    currCourse.course_code = cdata.course_code;
                    currCourse.course_id = cdata.id;
                    currCourse.ttl_stdnts = cdata.total_students;
                    currCourse.course_name = cdata.name.replace(/[^\s]+\s/i,""); //Extract course name only removing term, code and section (e.g.; Introduction to Accounting)
                    currCourse.enrollment_term_id = cdata.enrollment_term_id;
                    currCourse.quarter_name = /^\d{4}\D{2}/.exec(cdata.course_code); // Extract term name only (e.g. 2018WI)
                    currCourse.section = /(?<=SEC)\d+/.exec(cdata.course_code); // Extract section number only (e.g. 17)
                    currCourse.short_code = /(?<=\d\D\D_)\D+_\d+(?=-)/.exec(cdata.course_code); // Extract general course code only, removing term and section (e.g. Account_101)
                    if (controls.rptType == 'at-risk' || controls.rptType == 'participation') {
                        currCourse.discussions_due = 0;
                        currCourse.assignments_due = 0;
                    }
                    if (controls.rptType == 'instructor') {
                        currCourse.submissions_count = 0;
                        currCourse.submissions_graded_ontime = 0;
                        currCourse.submissions_graded_late = 0;
                        currCourse.submissions_graded_none = 0;
                        currCourse.graded_ontime_pcnt = 0;
                        currCourse.graded_late_pcnt = 0;
                        currCourse.graded_none_pcnt = 0;
                        currCourse.feedback_count = 0;
                        currCourse.feedback_mean_length = 0;
                    }
                    $('#dm_report_name').text(cdata.course_code);
                }
            }).fail(function() {
                var errorDetail = 'course,' + currCourse.course_id;
                throw new Error(errorDetail);
            }).done(function () {
                try {
                    if (currCourse.ttl_stdnts === 0) {
                        console.log(currCourse.course_code + " - No students enrolled.");
                        controls.emptyCourse = true;
                        controller('course_done');
                    } else {
                        var uURL = '/api/v1/courses/' + currCourse.course_id + '/users?&include[]=email';
                        uURL += controls.rptType == 'instructor' ? '&enrollment_type[]=teacher' : '';
                        uURL += '&per_page=100';
                        getUsers(uURL)
                    }
                } catch (e) {
                    errorHandler(e);
                }
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    function makeNewReport() {
        var currCourseID;
        try {
            if (!controls.combinedRpt) { // Clear userData object if outputting a separate report for each course
                for (var id in userData) {
                    delete userData[id];
                }
                accessData = [];
                atRiskArray = [];
                controls.accessIndex = 0;
                controls.accessCount = 0;
            }
            // Clear course-specific data from objects for single or multiple reports

            Object.keys(currCourse).forEach(function(key) { delete currCourse[key]; });
            Object.keys(enrollmentData).forEach(function(key) { delete enrollmentData[key]; });
            Object.keys(assignmentData).forEach(function(key) { delete assignmentData[key]; });
            Object.keys(topics).forEach(function(key) { delete topics[key]; });
            // clear topics info extracted from the current course
            if (controls.rptType == 'instructor') {
                topics = [];
                topicEntries = [];

                topicIDs = [];
                topicsIdx = 0;
            }
            if (controls.rptType == 'at-risk' || controls.rptType == 'instructor') {accessData = []}
            submissionData = [];
            controls.userIndex = -1;
            controls.userArray = [];
            accessData = [];
            controls.accessIndex = 0;
            if (controls.coursePending === 0) { // Output report if no more reports are pending
                outputReport();
            }
            aborted = false;
            controls.courseIndex = controls.courseArray.length - controls.coursePending;
            currCourseID = controls.courseArray[controls.courseIndex];
            progressbar(0,0); // Reset progress bar
            getCourseData(currCourseID);
        } catch (e) {
            errorHandler(e);
        }
    }

    // Get the Canvas course IDs that match the user's search criteria

    function getCourseIds(crsURL) {
        try {
            if (aborted) {
                throw new Error('Aborted');
            }
            $.getJSON(crsURL, function(cdata, status, jqXHR) {
                crsURL = nextURL(jqXHR.getResponseHeader('Link'));
                if (cdata) {
                    for (var i = 0; i < cdata.length; i++) {
                        var thisCourse = cdata[i];
                        var goodURL = true;
                        var crsNameStr = thisCourse.name;
                        var dlPttrn = /-DL_/;
                        var dlStatus = dlPttrn.test(crsNameStr); // Check if a course is online by the -DL_ pattern in link text
                        if ((controls.dlCrsOnly && dlStatus) || !controls.dlCrsOnly) { // Add courses to the list; either online only or all
                            var cncldPttrn = /_SECX\d\d/;
                            var cncldStatus = cncldPttrn.test(crsNameStr); // Check if current course is cancelled with "_SECX " in the course name
                            if (!cncldStatus) { // Add only courses without "_SECX " in their names
                                controls.courseArray.push(thisCourse.id);
                            }
                            if (cncldStatus) {
                                console.log(thisCourse.name + ' - Cancelled');
                                controls.emptyCourse = true;
                            }
                        }
                    }
                }
            }).done(function () {
                if (crsURL) {
                    getCourseIds(crsURL);
                } else {
                    controls.coursePending = controls.courseArray.length; //Count of courses to be processed
                    if (controls.coursePending === 0) {
                        alert('No courses matched your search criteria. Refine your search and try again.');
                        wrapup();
                        return false;
                    }
                    var runScriptDlg = 'The records from ' + controls.coursePending + ' courses will be processed. Continue?';
                    if (confirm(runScriptDlg) == true) {
                    makeNewReport(); // Get user accesses for each course selected
                    } else {
                        wrapup();
                    }
                }
            }).fail(function() {
                var errorDetail = 'courses IDs';
                throw new Error(errorDetail);
            });
        } catch (e) {
            errorHandler(e);
        }
    }

    // Processes user input from the report options dialog

    function setupReports(enrllTrm, srchBy, srchTrm) {
        var enrollTermID = (enrllTrm > 0) ? "&enrollment_term_id=" + enrllTrm : "";
        var searchBy = (srchBy == 'true') ? "" : "&search_by=teacher";
        var searchTrm = (srchTrm.length > 0) ? "&search_term=" + srchTrm : "";
        document.body.style.cursor = "wait";
        controls.emptyCourse = false;
        setupPool();
        progressbar();
        var cURL = '/api/v1/accounts/' + controls.canvasAcct + '/courses?with_enrollments=true' + enrollTermID + searchTrm + searchBy + '&per_page=100';
        getCourseIds(cURL);
    }


    // Report file generating functions below

    function outputReport() {
        var reportName = '';
        var rptDatesClean, rptDatesConcat;
        try {
            if (aborted) {
                console.log('Process aborted at makeReport');
                aborted = false;
                return;
            }
            $('#dm_report_status').text('Compiling report...');
            var csv = createCSV();
            if (csv) {
                var blob = new Blob([ csv ], {
                    'type' : 'text/csv;charset=utf-8'
                });
                if (controls.combinedRpt) {
                    reportName = $("#dm_term_slct option:selected").val() == 0 ? '' : $("#dm_term_slct option:selected").text() + ' ';
                    reportName += $('#dm_srch_inpt').val() == '' ? '' : $('#dm_srch_inpt').val().toUpperCase() + ' ';
                } else {
                    reportName = currCourse.course_code + ' ';
                }
                reportName += $("#dm_report_slct option:selected").text();
                if (controls.rptDateStartTxt != '') {
                    rptDatesConcat = controls.rptDateStartTxt + ' - ' + controls.rptDateEndTxt;
                    rptDatesClean = rptDatesConcat.replace(/\//g,'-'); // replace slashes with hyphens in dates for a vaild filename
                    reportName += ' ' + rptDatesClean.replace(/20(?=\d{2})/g,''); // remove century from dates
                }
                reportName +=  ' Report.csv';
                saveAs(blob, reportName);
                if (controls.coursePending > 0) {
                    makeNewReport();
                } else {
                    wrapup();
                }
            } else {
                throw new Error('Problem creating report');
            }
        } catch (e) {
            errorHandler(e);
        }
    }

    function createCSV() {
        var tatHrs;
        var fields
        if (controls.rptType == 'at-risk' || controls.rptType == 'participation') {
            fields = [ {
                'name' : 'User ID',
                'src' : 'u.id'
            }, {
                'name' : 'Login ID',
                'src' : 'u.login_id'
            }, {
                'name' : 'Email',
                'src' : 'u.email'
            }, {
                'name' : 'Total Hours Active',
                'src' : 'e.total_activity_time',
                'fmt' : 'hours'
            }, {
                'name' : 'Last Activity',
                'src' : 'e.last_activity_at',
                'fmt' : 'date'
            }, {
                'name' : 'Home Page Views',
                'src' : 'e.home_page_views'
            }, {
                'name' : 'Total Page Views',
                'src' : 'e.page_views'
            }, {
                'name' : 'Submitted / Due',
                'src' : 'e.submitted'
            }, {
                'name' : 'Late Assignments',
                'src' : 'e.assignments_late'
            }, {
                'name' : 'Max. Days Late',
                'src' : 'e.max_days_late',
                'fmt' : 'days'
            }, {
                'name' : 'Missing Assignments',
                'src' : 'e.assignments_missing'
            }, {
                'name' : 'Max. Days Missing',
                'src' : 'e.max_days_missing',
                'fmt' : 'days'
            }, {
                'name' : 'Current Score',
                'src' : 'e.current_score',
            }, {
                'name' : 'Discussion Posts / Due',
                'src' : 'e.discussion_posts'
            }, {
                'name' : 'Quarter',
                'src' : 'e.quarter_name',
            }, {
                'name' : 'Section',
                'src' : 'e.section',
            }, {
                'name' : 'Short Course Code',
                'src' : 'e.short_course_code',
            }, {
                'name' : 'Course Name',
                'src' : 'e.course_name',
            }, {
                'name' : 'Full Course Code',
                'src' : 'e.course_code',
            }, {
                'name' : 'Instructor Name',
                'src' : 'e.teacher_name'
            }, {
                'name' : 'Instructor Email',
                'src' : 'e.teacher_email'
            }, {
                'name' : 'Student Course Enrollment Page',
                'src' : 'e.html_url'
            }];
        } else if (controls.rptType == 'instructor') {
            fields = [ {
                'name' : 'User ID',
                'src' : 'u.id'
            }, {
                'name' : 'Login ID',
                'src' : 'u.login_id'
            }, {
                'name' : 'Email',
                'src' : 'u.email'
            }, {
                'name' : 'Total Hours Active',
                'src' : 'e.total_activity_time',
                'fmt' : 'hours'
            }, {
                'name' : 'Last Activity',
                'src' : 'a.last_activity_at',
                'fmt' : 'date'
            }, {
                'name' : 'Home Page Views',
                'src' : 'a.home_page_views'
            }, {
                'name' : 'Total Page Views',
                'src' : 'a.page_views'
            }, {
                'name' : 'Discussion Posts',
                'src' : 'e.discussion_posts'
            }, {
                'name' : 'Last Post Date',
                'src' : 'e.last_post_date',
                'fmt' : 'date'
            }, {
                'name' : 'Mean Post Chars',
                'src' : 'e.post_mean_length',
                'fmt' : 'integer',
            }, {
                'name' : 'Graded On-time %',
                'src' : 'e.graded_ontime_pcnt',
            }, {
                'name' : 'Graded Late %',
                'src' : 'e.graded_late_pcnt',
            }, {
                'name' : 'Grades Overdue %',
                'src' : 'e.graded_none_pcnt',
            }, {
                'name' : 'Assignment Feedback',
                'src' : 'e.feedback_count',
            }, {
                'name' : 'Mean Feedback Chars',
                'src' : 'e.feedback_mean_length',
                'fmt' : 'integer',
            }, {
                'name' : 'Quarter',
                'src' : 'a.quarter_name',
            }, {
                'name' : 'Enrollment',
                'src' : 'a.ttl_stdnts',
            }, {
                'name' : 'Section',
                'src' : 'a.section',
            }, {
                'name' : 'Short Course Code',
                'src' : 'a.short_course_code',
            }, {
                'name' : 'Course Name',
                'src' : 'a.course_name',
            }, {
                'name' : 'Full Course Code',
                'src' : 'a.course_code',
            }, {
                'name' : 'Instructor Course Enrollment Page',
                'src' : 'a.html_url'
            }];
        } else {
            fields = [{
                'name' : 'User ID',
                'src' : 'u.id'
            }, {
                'name' : 'Login ID',
                'src' : 'u.login_id'
            }, {
                'name' : 'Role',
                'src' : 'a.user_role'
            }, {
                'name' : 'Total Hours Active',
                'src' : 'a.total_activity_time',
                'fmt' : 'hours'
            }, {
                'name' : 'Asset Title',
                'src' : 'a.readable_name'
            }, {
                'name' : 'Views',
                'src' : 'a.view_score'
            }, {
                'name' : 'Participations',
                'src' : 'a.participate_score'
            }, {
                'name' : 'First Access',
                'src' : 'a.created_at',
                'fmt' : 'date'
            }, {
                'name' : 'Last Access',
                'src' : 'a.last_access',
                'fmt' : 'date'
            }, {
                'name' : 'Action',
                'src' : 'a.action_level'
            }, {
                'name' : 'Asset Code',
                'src' : 'a.asset_code'
            }, {
                'name' : 'Asset Group Code',
                'src' : 'a.asset_group_code'
            }, {
                'name' : 'Quarter',
                'src' : 'a.quarter_name',
            }, {
                'name' : 'Section',
                'src' : 'a.section',
            }, {
                'name' : 'Short Course Code',
                'src' : 'a.short_course_code',
            }, {
                'name' : 'Course Name',
                'src' : 'a.course_name',
            }, {
                'name' : 'Full Course Code',
                'src' : 'a.course_code',
            }, {
                'name' : 'Canvas Course ID',
                'src' : 'a.course_id',
            }, {
                'name' : 'SIS Course ID',
                'src' : 'a.sis_course_id',
            }, {
                'name' : 'Canvas Term ID',
                'src' : 'a.enrollment_term_id',
            }, {
                'name' : 'Asset Category',
                'src' : 'a.asset_category'
            }, {
                'name' : 'Asset Class',
                'src' : 'a.asset_class_name'
            }];
        }
        if (!controls.anonStdnts) {
            fields.splice(2, 0, {
                'name' : 'Sortable Name',
                'src' : 'u.sortable_name'
            });
        } else {
            fields.splice(0, 2, {
                'name' : 'Name',
                'src' : 'u.name'
            });
        }
        if (controls.rptType == 'participation') {
           //fields.splice(6, 10);
            //fields.splice(4, 1);
        }
        var canSIS = false;
        for (var id in userData) {
            if (userData.hasOwnProperty(id)) {
                if (typeof userData[id].sis_user_id !== 'undefined' && userData[id].sis_user_id) {
                    canSIS = true;
                    break;
                }
            }
        }
        var CRLF = '\r\n';
        var hdr = [];
        fields.map(function(e) {
            if (typeof e.sis === 'undefined' || (e.sis && canSIS)) {
                var name = (typeof headingNoSpaces !== 'undefined' && headingNoSpaces) ? e.name.replace(' ', '') : e.name;
                hdr.push(name);
            }
        });
        var t = hdr.join(',') + CRLF;
        var item, user, enrollment, userId, fieldInfo, value, currArray;
        switch (controls.rptType) {
            case 'access':
                currArray = accessData;
                break;
            case 'instructor':
                currArray = instructorData;
                break;
            default:
                currArray = atRiskArray;
                }
        for (var i = 0; i < currArray.length; i++) {
            item = currArray[i];
            if (controls.rptType == 'access' && (typeof (showViewStudent) === 'undefined' || !showViewStudent) && item.asset_category == 'roster' && item.asset_class_name == 'student_enrollment') {
                continue;
            }
            userId = item.user_id;
            user = userData[userId];
            if (typeof (user) === 'undefined' || !user) {continue}
            for (var j = 0; j < fields.length; j++) {
                if (typeof fields[j].sis !== 'undefined' && fields[j].sis && !canSIS) {
                    continue;
                }
                fieldInfo = fields[j].src.split('.');
                switch(fieldInfo[0]) {
                    case 'a':
                        value = item[fieldInfo[1]];
                        if (fieldInfo[1] == 'total_activity_time') {
                            if (item.readable_name == 'Course Home') {
                                value = item[fieldInfo[1]];
                            } else {
                                value = '';
                            }
                        } else {
                            value = item[fieldInfo[1]];
                        }
                        break;
                    case 'e':
                        value = item[fieldInfo[1]];
                        break;
                    default:
                        if (controls.anonStdnts && fieldInfo[1] == 'name' && (item.user_role == 'Student' || item.role == 'Student')) {
                            value = user.anon_name;
                        } else {
                            value = user[fieldInfo[1]];
                        }
                }
                if (typeof value === 'undefined' || value === null) {
                    value = '';
                } else {
                    if (typeof fields[j].fmt !== 'undefined') {
                        switch (fields[j].fmt) {
                            case 'date':
                                value = excelDate(value);
                                break;
                            case 'hours':
                                value = timeHours(value);
                                break;
                            case 'days':
                                value = timeDays(value);
                                break;
                            case 'integer':
                                value = integerVal(value);
                                break;
                            default:
                                break;
                        }
                    }
                    if (typeof value === 'string') {
                        var quote = false;
                        if (value.indexOf('"') > -1) {
                            value = value.replace(/"/g, '""');
                            quote = true;
                        }
                        if (value.indexOf(',') > -1) {
                            quote = true;
                        }
                        if (quote) {
                            value = '"' + value + '"';
                        }
                    }
                }
                if (j > 0) {
                    t += ',';
                }
                t += value;
            }
            t += CRLF;
        }
        return t;
    }

    function excelDate(timestamp) {
        var d;
        try {
            if (!timestamp) {
                return '';
            }
            timestamp = timestamp.replace('Z', '.000Z');
            var dt = new Date(timestamp);
            if (typeof dt !== 'object') {
                return '';
            }
            d = dt.getFullYear() + '-' + pad(1 + dt.getMonth()) + '-' + pad(dt.getDate()) + ' ' + pad(dt.getHours()) + ':' + pad(dt.getMinutes()) + ':' + pad(dt.getSeconds());
        } catch (e) {
            errorHandler(e);
        }
        return d;

        function pad(n) {
            return n < 10 ? '0' + n : n;
        }
    }

    function timeDays(seconds) {
        var d, dRaw;
        try {
            if (!seconds || isNaN(seconds)) {
                return '';
            }
            dRaw = seconds / 86400;
            d = dRaw.toFixed(2);
        } catch (e) {
            errorHandler(e);
        }
        return d;
    }

    function timeHours(seconds) {
        var h, hRaw;
        try {
            if (!seconds || isNaN(seconds)) {
                return '';
            }
            hRaw = seconds / 3600;
            h = hRaw.toFixed(2);
        } catch (e) {
            errorHandler(e);
        }
        return h;
    }

    function integerVal(float) {
        var i;
        try {
            if (!float || isNaN(float)) {
                return '';
            }
            i = Math.round(float);
        } catch (e) {
            errorHandler(e);
        }
        return i;
    }

    // User interface functions below

    // Clear or reset objects, arrays and vars of all data, reset UI
    function wrapup() {
        progressbar(); // Close progress bar
        Object.keys(userData).forEach(function(key) { delete userData[key]; });
        Object.keys(currCourse).forEach(function(key) { delete currCourse[key]; });
        Object.keys(enrollmentData).forEach(function(key) { delete enrollmentData[key]; });
        Object.keys(assignmentData).forEach(function(key) { delete assignmentData[key]; });
        submissionData = [];
        accessData = [];
        atRiskArray = [];
        aborted = false;
        controls.accessCount = 0;
        controls.accessIndex = 0;
        controls.courseArray = [];
        controls.courseIndex = 0;
        controls.coursePending = -1;
        controls.userArray = [];
        controls.userIndex = -1
        tchrNameArray.length = 0;
        tchrEmailArray.length = 0;
        controls.rptDateStart = 0;
        controls.rptDateEnd = 0;
        controls.rptDateStartTxt = '';
        controls.rptDateEndTxt = '';
        document.body.style.cursor = "default"; // Restore default cursor
        $('#jj_access_report').one('click', reportOptionsDlg); // Re-enable multi-reports button
        if (controls.emptyCourse) {
            alert('Courses labeled as cancelled (SECX##) and those with no students enrolled have been omitted. See Console for a list of omitted courses.');
        }
    }

    function progressbar(total, current) {
        try {
            if (typeof total === 'undefined' || typeof current == 'undefined') {
                if ($('#jj_progress_dialog').length === 0) {
                    $('body').append('<div id="jj_progress_dialog"></div>');
                    $('#jj_progress_dialog').append('<div id="dm_report_name" style="font-size: 12pt; font-weight:bold"></div>');
                    $('#jj_progress_dialog').append('<div id="jj_progressbar"></div>');
                    $('#jj_progress_dialog').append('<div id="dm_report_status" style="font-size: 12pt; text-align: center"></div>');
                    $('#jj_progress_dialog').dialog({
                        'title' : 'Fetching Canvas Data',
                        'autoOpen' : false,
                        'buttons' : [ {
                            'text' : 'Cancel',
                            'click' : function() {
                                $(this).dialog('close');
                                aborted = true;
                                abortAll();
                                $('#jj_access_report').one('click', reportOptionsDlg);
                            }
                        } ]
                    });
                }
                if ($('#jj_progress_dialog').dialog('isOpen')) {
                    $('#jj_progress_dialog').dialog('close');
                } else {
                    $('#jj_progressbar').progressbar({
                        'value' : false
                    });
                    $('#jj_progress_dialog').dialog('open');
                }
            } else {
                if (!aborted) {
                    var val = current > 0 ? Math.round(100 * total / current) : false;
                    $('#jj_progressbar').progressbar('option', 'value', val);
                }
            }
        } catch (e) {
            errorHandler(e);
        }
    }

    function dateOptionsDlgResult(result) {
        var dS, hS, dE, hE;
        if (result) {
            controls.rptDateStartTxt = $("#dm_start_date_txt").val();
            controls.rptDateEndTxt = $("#dm_end_date_txt").val();
            dS = Date.parse($("#dm_start_date_txt").val());
            hS = dS.setHours(0,0,0,0);
            controls.rptDateStart = hS;
            dE = Date.parse($("#dm_end_date_txt").val());
            hE = dE.setHours(24,0,0,0);
            controls.rptDateEnd = hE;
            $('#dm_date_range_p').html(controls.rptDateStartTxt + " - " + controls.rptDateEndTxt);
        } else {
            controls.rptDateStart = 0;
            controls.rptDateEnd = 0;
            controls.rptDateStartTxt = '';
            controls.rptDateEndTxt = '';
            $("#dm_rprt_prd_chbx").prop("checked",false);
            $('#dm_date_range_p').html('');
        }
    }

    function showDateOptionsDlg() {
        try {
            if ($('#dm_prtcptn_dts_dialog').length === 0) {
                $('body').append('<div id="dm_prtcptn_dts_dialog"></div>');
                $('#dm_prtcptn_dts_dialog').append('<p style="font-size: 1em">Select range of dates for zero participation.</p>');
                $('#dm_prtcptn_dts_dialog').append('<form id="dm_prtcptn_dts_frm"></div>');
                $('#dm_prtcptn_dts_frm').append('<fieldset id="dm_prtcptn_dts_fldst"></fieldset>');
                $('#dm_prtcptn_dts_fldst').append('<label for="dm_start_date_txt">Start Date:</label>');
                $('#dm_prtcptn_dts_fldst').append('<input type="text" id="dm_start_date_txt" name="dm_start_date_txt">');
                $('#dm_prtcptn_dts_fldst').append('<label for="dm_end_date_txt">End Date:</label>');
                $('#dm_prtcptn_dts_fldst').append('<input type="text" id="dm_end_date_txt" name="dm_end_date_txt">');
                var enableDataPicker = $( function() {
                    $( "#dm_start_date_txt" ).datepicker();
                    $( "#dm_end_date_txt" ).datepicker();
                } );
                $('#dm_prtcptn_dts_dialog').dialog ({
                    'title' : 'Specify Reporting Period',
                    'autoOpen' : false,
                    buttons : {
                        "OK": function () {
                            var dS = Date.parse($("#dm_start_date_txt").val());
                            var dE = Date.parse($( "#dm_end_date_txt").val());
                            if (!(dS instanceof Date) || !(dE instanceof Date)) {
                                alert('Enter valid start and end dates for the reporting period')
                                return false;
                            } else if (dE < dS) {
                                alert('The end date must occur after the start date.');
                                return false;
                            } else {
                                $(this).dialog("close");
                                dateOptionsDlgResult(true);
                            }
                        },
                        "Cancel": function () {
                            $(this).dialog("close");
                            dateOptionsDlgResult(false);
                        }
                    }});
            }
            $('#dm_prtcptn_dts_dialog').dialog('open');
        } catch (e) {
            return false;
            errorHandler(e);
        }
    }

    function enableReportOptionsDlgOK() {
        if (($("#dm_term_slct").val() != 0 || $("#dm_srch_inpt").val() != '') && $("#dm_report_slct").val() != 0) {
            $('#dm_term_slct').closest(".ui-dialog").find("button:contains('OK')").removeAttr('disabled').removeClass( 'ui-state-disabled' );;
        } else {
             $('#dm_term_slct').closest(".ui-dialog").find("button:contains('OK')").prop("disabled", true).addClass("ui-state-disabled");
        }
    }

    function reportOptionsDlg() {
        try {
            if ($('#dm_options_frm').length === 0) {
                // Update this array with new Canvas term IDs and labels as quarters/terms are added
                // Populates the term select menu in the "Select Report Options" dialog box
                var terms = {data:[
                    {val : 0, txt: 'Select a term'},
                    {val : 166, txt: '2020 Summer'},
                    {val : 165, txt: '2020 Spring'},
                    {val : 164, txt: '2020 Winter'},
                    {val : 163, txt: '2019 Fall'},
                    {val : 129, txt: '2019-2020 Academic Year'},
                    {val : 131, txt: '2019-2020 Academic Year'},
                    {val : 128, txt: '2019 Summer'},
                    {val : 127, txt: '2019 Spring'},
                    {val : 124, txt: '2019 Winter'},
                    {val : 126, txt: '2018 Fall'},
                    {val : 125, txt: '2018-2019 Med Academic Year'},
                    {val : 130, txt: '2018-2019 Academic Year'},
                    {val : 123, txt: '2018 Summer'},
                    {val : 122, txt: '2018 Spring'},
                    {val : 121, txt: '2018 Winter'},
                    {val : 120, txt: '2017 Fall'},
                    {val : 118, txt: '2017-2018 Academic Year'},
                    {val : 119, txt: '2017 Summer'},
                    {val : 113, txt: '2017 Spring'},
                    {val : 112, txt: '2017 Winter'},
                    {val : 111, txt: '2016 Fall'},
                    {val : 109, txt: '2016-2017 Academic Year'},
                    {val : 110, txt: '2016 Summer'},
                    {val : 107, txt: '2016 Spring'},
                    {val : 106, txt: '2016 Winter'},
                    {val : 105, txt: '2015 Fall'},
                    {val : 108, txt: '2015-2016 Academic Year'},
                    {val : 103, txt: '2015 Summer'},
                    {val : 93, txt: '2015 Spring'},
                    {val : 96, txt: '2015 Winter'},
                    {val : 92, txt: '2014 Fall'},
                    {val : 104, txt: '2014-2015 Academic Year'},
                    {val : 115, txt: 'Advising Term'},
                    {val : 1, txt: 'Default Term'},
                    {val : 116, txt: 'Demo Term'},
                    {val : 114, txt: 'Prep Site Term'},
                    {val : 117, txt: 'Program Term'}
                ]};
                // Populates the reports select menu in the "Select Report Options" dialog box
                var reports = {data:[
                    {val : '0', txt: 'Select a report type'},
                    {val : 'at-risk', txt: 'At-risk Students'},
                    {val : 'access', txt: 'Course Resource Access'},
                    {val : 'instructor', txt: 'Instructor Presence'},
                    {val : 'participation', txt: 'Zero Participation'},
                ]};
                // Define "Select Report Options" dialog box
                $('body').append('<div id="dm_options_dialog"></div>');
                $('#dm_options_dialog').append('<form id="dm_options_frm"></div>');
                $('#dm_options_frm').append('<fieldset id="dm_options_fldst"></fieldset>');
                $('#dm_options_fldst').append('<select id="dm_term_slct">');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<input type="radio" name="dm_srch_rdo" id="coursename" value="true" checked="checked">');
                $('#dm_options_fldst').append('<label for="coursename">&nbsp;Course Name</label>');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<input type="radio" name="dm_srch_rdo" id="instructorname" value="false">');
                $('#dm_options_fldst').append('<label for="instructorname">&nbsp;Instructor Name</label>');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<label for="dm_srch_inpt">Search text:</label>');
                $('#dm_options_fldst').append('<input type="text" id="dm_srch_inpt" name="dm_srch_inpt">');
                $('#dm_options_fldst').append('<hr/>');
                $('#dm_options_fldst').append('<select id="dm_report_slct">');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<input type="radio" name="dm_report_opts" id="single" value="true" checked="checked">');
                $('#dm_options_fldst').append('<label for="single">&nbsp;Single report (all selected courses)</label>');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<input type="radio" name="dm_report_opts" id="multiple" value="false">');
                $('#dm_options_fldst').append('<label for="multiple">&nbsp;Multiple reports (individual courses)</label>');
                $('#dm_options_fldst').append('<hr/>');
                $('#dm_options_fldst').append('<input type="checkbox" id="dm_dl_crs_chbx" name="dm_dl_crs_chbx" value="true" checked>');
                $('#dm_options_fldst').append('<label for="dm_dl_crs_chbx">&nbsp;Online courses only</label>');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<input type="checkbox" id="dm_anon_stdnt_chbx" name="dm_anon_stdnt_chbx" value="true" checked>');
                $('#dm_options_fldst').append('<label for="dm_anon_stdnt_chbx">&nbsp;Anonymize students</label>');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<input type="checkbox" id="dm_rprt_prd_chbx" name="dm_rprt_prd_chbx" value="true">');
                $('#dm_options_fldst').append('<label for="dm_rprt_prd_chbx">&nbsp;Specify reporting period</label>');
                $('#dm_options_fldst').append('<br/>');
                $('#dm_options_fldst').append('<span id="dm_date_range_p" style="font-size: 1em; margin-left: 1.75em"></span>');
                $("#dm_term_slct").change(function() {
                    enableReportOptionsDlgOK();
                });
                $("#dm_srch_inpt").change(function() {
                    enableReportOptionsDlgOK();
                });
                $("#dm_report_slct").change(function() {
                    if ($(this).children("option:selected").val() == 'participation')
                    {
                        $("#dm_anon_stdnt_chbx").prop("checked",false);
                        $("#dm_anon_stdnt_chbx").prop("disabled",true);
                        $("#dm_rprt_prd_chbx").prop("checked",true);
                        $("#dm_rprt_prd_chbx").prop("disabled",true);
                        showDateOptionsDlg();
                    }
                    else if ($(this).children("option:selected").val() == 'at-risk' || $(this).children("option:selected").val() == 'instructor')
                    {
                        $("#dm_anon_stdnt_chbx").prop("checked",false);
                        $("#dm_anon_stdnt_chbx").prop("disabled",true);
                         $("#dm_rprt_prd_chbx").prop("disabled",false);
                    }
                    else {
                        controls.rptDateStart = null;
                        controls.rptDateEnd = null;
                        $("#dm_anon_stdnt_chbx").prop("disabled",false);
                        $("#dm_rprt_prd_chbx").prop("disabled",false);
                        $("#dm_rprt_prd_chbx").prop("checked",false);
                        $('#dm_date_range_p').html('');
                    }
                    enableReportOptionsDlgOK();
                });
                $("#dm_rprt_prd_chbx").change(function() {
                    controls.rptDateStart = 0;
                    controls.rptDateEnd = 0;
                    if($(this).is(":checked"))
                    {
                        showDateOptionsDlg();
                    } else {
                        $('#dm_date_range_p').html('');
                    }
                });
                $('#dm_options_dialog').dialog ({
                    'title' : 'Select Report Options',
                    'autoOpen' : false,
                    'modal' : true,
                    'buttons' : {
                        "OK": function () {
                            $(this).dialog("close");
                            var enrllTrmSelct = $("#dm_term_slct option:selected").val();
                            var srchByChecked = $("input[name='dm_srch_rdo']:checked").val();
                            var srchTermsStr = $("#dm_srch_inpt").val();
                            controls.rptType = $('#dm_report_slct').children("option:selected").val();
                            controls.combinedRpt = $("input[name='dm_report_opts']:checked").val()
                            controls.dlCrsOnly = $('#dm_dl_crs_chbx').prop('checked');
                            controls.anonStdnts = $('#dm_anon_stdnt_chbx').prop('checked');
                            setupReports(enrllTrmSelct, srchByChecked, srchTermsStr);
                        },
                        "Cancel": function () {
                            $(this).dialog('close');
                            $('#jj_access_report').one('click', reportOptionsDlg);
                        }
                    }});
                $('.ui-dialog-titlebar-close').remove(); // Remove titlebar close button forcing users to form buttons
                $('#dm_term_slct').closest(".ui-dialog").find("button:contains('OK')").prop("disabled", true).addClass("ui-state-disabled");
                if ($('#dm_term_slct').children('option').length === 0) { // add terms to terms select element
                    $.each(terms.data, function (key, value) {
                        $("#dm_term_slct").append($('<option>', {
                            value: value.val,
                            text: value.txt,
                            'data-mark': value.id
                        }));
                    });
                }
                if ($('#dm_report_slct').children('option').length === 0) { // add report types to reports select element
                    $.each(reports.data, function (key, value) {
                        $("#dm_report_slct").append($('<option>', {
                            value: value.val,
                            text: value.txt,
                            'data-mark': value.id
                        }));
                    });
                }
            }
            //$('#dm_options_dialog').css('z-index', '9000');
            $('#dm_options_dialog').dialog('open');
        } catch (e) {
            errorHandler(e);
        }
    }

    // Add "Multi-Course Access Reports" button below navigation
    function addReportsLink() {
        if ($('#jj_access_report').length === 0) {
            $('#left-side').append('<div class="rs-margin-bottom"><a id="jj_access_report"><span aria-hidden="true" style="color:#f92626; cursor: pointer; display: block; font-size: 1rem; line-height: 20px; margin: 5px auto; padding: 8px 0px 8px 6px;">Custom API Reports</span><span class="screenreader-only">Custom API Reports</span></a></div>');
            $('#jj_access_report').one('click', reportOptionsDlg);
        }
        return;
    }

    $( document ).ready(function() {
        addReportsLink(); // Add reports link to page
    });

}());
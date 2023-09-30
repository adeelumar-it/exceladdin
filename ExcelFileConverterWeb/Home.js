// <reference path="messageread.js" />
var app = angular.module('edgelegal', ['ngMaterial', "ngRoute"], function () {


});






app.controller('edgelegalctrl', function ($scope, $mdDialog, $mdToast, $log, $location,) {


           
    
    $scope.ShowMainDiv = false;
    var filecontent = "";


    Office.onReady(function (info) {

        if (info.host === Office.HostType.Excel) {

            ProgressLinearActive();
            let dialog; // Declare dialog as global for use in later functions.
            let MatterNumberdialog;



            let userInfo = window.localStorage.getItem('userinfo');
            userInfo = JSON.parse(userInfo)
            if (userInfo) {

                if (userInfo.userName) {
                    $scope.ShowMainDiv = true;
                    $scope.userName = userInfo.userName
                    ProgressLinearInActive();
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }


                } else {

                    $scope.ShowMainDiv = false;
                    ProgressLinearInActive();
                    openDialog();

                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }

            } else {
                $scope.ShowMainDiv = false;
                ProgressLinearInActive();

                openDialog();

            }
            function openDialog() {
                Office.context.ui.displayDialogAsync('https://localhost:44311/Templates/Login.html', { height: 50, width: 30 },
                    function (asyncResult) {
                        dialog = asyncResult.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                    }
                );
            }



            function processMessage(arg) {
                console.log(arg)
                let message = JSON.parse(arg.message);
                window.localStorage.setItem('userinfo', JSON.stringify(message));
                let userdata = message;
                message = message.login;

                //dialog.close();

                if (message === true) {
                    console.log(message)
                    $scope.userName = userdata.userName
                    // Authentication was successful, close the dialog, and perform other actions as needed
                    dialog.close();
                    console.log("log in")
                    $scope.ShowMainDiv = true;
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                    //$scope.Message = false;
                } else {
                    dialog.close();
                    $scope.ShowMainDiv = true;
                    //$scope.Message = true;
                }
            }
            $scope.getFilebase64 = function (ev) {

                Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 },
                    function (result) {
                        if (result.status === "succeeded") {
                            const myFile = result.value;
                            myFile.getSliceAsync(0, function (result) {

                                let data = result.value.data

                                let btoadata = btoa(String.fromCharCode.apply(null, new Uint8Array(data)))



                                if (btoadata) {
                                    myFile.closeAsync();
                                    filecontent = btoadata;
                                    Excel.run(function (context) {
                                        var workbook = context.workbook;

                                        workbook.load(["name"]);

                                        return context.sync()
                                            .then(function () {
                                                // Access the workbook name
                                                var workbookName = workbook.name;
                                                $scope.Filename = workbookName;
                                                Office.context.ui.displayDialogAsync(`https://localhost:44311/Templates/MatterNumber.html?workbookName=${workbookName}`, { height: 50, width: 30 },
                                                    function (asyncResult) {
                                                        MatterNumberdialog = asyncResult.value;
                                                        MatterNumberdialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                                                    }
                                                );

                                                //let fileobj = {
                                                //    fileName: $scope.Filename,

                                                //    fileContent: fileContent,

                                                //}
                                                //window.localStorage.setItem("filedata", JSON.stringify(fileobj))



                                                //ProgressLinearInActive();
                                                console.log("Active workbook name: " + workbookName);
                                            })
                                            .catch(function (error) {
                                                console.log("Error: " + error);
                                            });
                                    }).catch(function (error) {
                                        console.log("Error: " + error);
                                    });
                                }



                            });

                        } else {
                            myFile.closeAsync()

                            // Handle the error here
                            loadToast("Upload Error");
                            //ProgressLinearInActive();

                        }


                    }
                );








                //ProgressLinearActive();
                //$mdDialog.show({

                //    scope: $scope.$new(),
                //    templateUrl: 'Templates/MatterNumber.html',
                //    targetEvent: ev,
                //    fullscreen: $scope.customFullscreen,
                //    controller: ['$scope', '$mdDialog', function ($scope, $mdDialog) {


                //        $scope.Filename = "";
                //        var fileContent = "";


                //        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 },
                //            function (result) {
                //                if (result.status === "succeeded") {
                //                    const myFile = result.value;
                //                    myFile.getSliceAsync(0, function (result) {

                //                        let data = result.value.data

                //                        let btoadata = btoa(String.fromCharCode.apply(null, new Uint8Array(data)))



                //                        if (btoadata) {
                //                            myFile.closeAsync();
                //                            fileContent = btoadata;
                //                            Excel.run(function (context) {
                //                                var workbook = context.workbook;
                //                                //var activeSheet = workbook.getActiveWorksheet();

                //                                // Load the workbook properties, including the name
                //                                workbook.load(["name"]);

                //                                return context.sync()
                //                                    .then(function () {
                //                                        // Access the workbook name
                //                                        var workbookName = workbook.name;
                //                                       $scope.Filename = workbookName;

                //                                         //let fileobj = {

                //                                         // matterNumber: $scope.MatterNumber,
                //                                         // fileContent: btoadata

                //                                         //}


                //                                        ProgressLinearInActive();
                //                                        console.log("Active workbook name: " + workbookName);
                //                                    })
                //                                    .catch(function (error) {
                //                                        console.log("Error: " + error);
                //                                    });
                //                            }).catch(function (error) {
                //                                console.log("Error: " + error);
                //                            });
                //                        }



                //                    });

                //                } else {
                //                    myFile.closeAsync()

                //                    // Handle the error here
                //                    loadToast("Upload Error");
                //                    ProgressLinearInActive();

                //                }


                //            }
                //        );

                //        $scope.Upload = function () {
                //            let fileobj = {
                //                fileName: $scope.Filename,
                //                matterNumber: $scope.MatterNumber,
                //                fileContent: fileContent,

                //            }

                //            console.log(fileobj)
                //            $mdDialog.hide();

                //        }
                //         $scope.Cancel = function () {

                //             $mdDialog.hide();
                //        }



                //    }]
                //});

            }
            function processMessage(arg) {

                console.log(arg)
                let message = JSON.parse(arg.message);
                if (message.close == true) {

                    MatterNumberdialog.close();
                    ProgressLinearInActive();

                } else {
                    message.filecontent = filecontent;
                    console.log(message)
                    ProgressLinearInActive();
                    MatterNumberdialog.close();

                }




            }


            $scope.Logout = function () {
                window.localStorage.clear('userinfo');
                window.location.reload();

            }

            $scope.Help = function () {
                window.open("https://support.microsoft.com/en-us")

            }


        } else {


            loadtost("word addin is wroking ")
            loadtost("no functioanlities included yet")

        }
      



        //$scope.getbase = function () {

        //    var settings = {
        //        "url": "http://223.27.18.219:9080/LPDM/RT/WS/addTemplate",
        //        "method": "POST",
        //        "timeout": 0,
        //        "headers": {
        //            "Content-Type": "application/json"
        //        },
        //        "data": JSON.stringify("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"),
        //    };

        //    $.ajax(settings).done(function (response) {
        //        console.log(response);
        //    });
        //    //Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 },
        //    //    function (result) {
        //    //        if (result.status == "succeeded") {
        //    //            const myFile = result.value;
        //    //            const sliceCount = myFile.sliceCount;
        //    //            const docdataSlices = [];
        //    //            let slicesReceived = 0,
        //    //                gotAllSlices = true;

        //    //            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
        //    //        } else {
        //    //            // Handle the error here
        //    //        }
        //    //});
        //}

        //function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
        //    file.getSliceAsync(nextSlice, function (sliceResult) {
        //        if (sliceResult.status == "succeeded") {
        //            if (!gotAllSlices) {
        //                return;
        //            }

        //            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
        //            if (++slicesReceived == sliceCount) {
        //                file.closeAsync();
        //                onGotAllSlices(docdataSlices);
        //            } else {
        //                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
        //            }
        //        } else {
        //            gotAllSlices = false;
        //            file.closeAsync();
        //            // Handle the error here
        //        }
        //    });
        //}

        //function onGotAllSlices(gotAllSlices) {
        //    let docdata = [];
        //    for (let i = 0; i < gotAllSlices.length; i++) {
        //        docdata = docdata.concat(gotAllSlices[i]);
        //    }

        //    let fileContent = "";
        //    for (let j = 0; j < docdata.length; j++) {
        //        fileContent += String.fromCharCode(docdata[j]);
        //    }

        //    const base64Data = btoa(fileContent);

        //    console.log(base64Data);

        //    // Now 'base64Data' contains the Base64 representation of the file content.
        //    // You can use it as needed.

        //    // For example, if you want to send it to a server:
        //    // sendDataToServer(base64Data);
        //}

        // Add any other necessary code here

    });



   
       


        function ProgressLinearActive() {
            $("#StartProgressLinear").show(function () {

                $("#ProgressBgDiv").show();
                $scope.ddeterminateValue = 15;
                $scope.showProgressLinear = false;
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            });
        };
        function ProgressLinearInActive() {
            $("#StartProgressLinear").hide(function () {
                setTimeout(function () {
                    $scope.ddeterminateValue = 0;
                    $scope.showProgressLinear = true;
                    $("#ProgressBgDiv").hide();
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }, 500);
            });
        };
        function loadToast(alertMessage) {
            var el = document.querySelectorAll('#zoom');
            $mdToast.show(
                $mdToast.simple()
                    .textContent(alertMessage)
                    .position('bottom')
                    .hideDelay(4000))
                .then(function () {
                    $log.log('Toast dismissed.');
                }).catch(function () {
                    $log.log('Toast failed or was forced to close early by another toast.');
                });
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        };

        if (!$scope.$$phase) {
            $scope.$apply();
        }

  
      
   
})

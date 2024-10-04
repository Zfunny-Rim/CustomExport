define([
    "qlik",
    "jquery",
    "text!./template.html",
    "css!./styles.css",
    "./xlsx.full.min",
], function (qlik, $, templateHtml, contentCss, XLSX) {
    var objectMap = objectMap || new Map();  // 전역 변수 유지
    var extId = '';
    return {
        // 사용자 정의 속성 패널 정의 (definition)
        definition: {
            type: "items",
            component: "accordion",
            items: {
                settings: {
                    uses: "settings",
                    items: {
                        serverAddress: {
                            ref: "serverAddress",
                            label: "Server Address for Encryption",
                            type: "string",
                            expression: "optional",
                            defaultValue: "https://your-server-address.com"
                        }
                    }
                }
            }
        },

        // paint 함수: 기존 코드 유지
        paint: function ($element, layout) {
            $element.html(templateHtml);
            // 시트 위에 팝업 추가하기 (시트 영역에 동적 팝업 삽입)
            extId = layout.qInfo.qId;

            if (!$("#globalExportPopup").length) {
                $("body").append(`
                    <div id="overlay"></div>  <!-- 배경 덮는 레이어 -->
                    <div id="globalExportPopup" class="popupModal" style="display:none;">
                        <h3>Select Objects to Export</h3>
                        <!--<div id="objectSelectContainer"></div>-->
                        <table id="objectSelectContainer">
                            <thead>
                                <tr>
                                    <td></td>
                                    <td>TYPE</td>
                                    <td>TITLE</td>
                                    <td>ID</td>
                                </tr>
                            </thead>
                            <tbody id="objectSelectTbody"></tbody>
                        </table>
                        <button id="selectAll">Select All</button>
                        
                        <div>
                            <input type="checkbox" id="encryptCheckbox"> <label for="encryptCheckbox">Enable Encryption</label>
                            <div id="passwordInput" style="display: none;">
                                <label for="password">Password:</label>
                                <input type="password" id="password" style="width: 100%;">
                            </div>
                        </div>
                        
                        <button id="exportBtn">Export</button>
                        <button id="closeModalBtn">Close</button>
                    </div>
                    `);
            }

            // Excel Export 버튼 클릭 시
            $("#excelExport").click(function () {
                $("#globalExportPopup").show();
                $("#overlay").show();  // 백그라운드 덮기 활성화
                populateExportableObjects(); // 오브젝트 목록을 드롭다운에 채움
            });

            // 팝업 닫기 버튼
            $("#closeModalBtn, #overlay").click(function () {
                $("#encryptCheckbox").prop('checked',false);
                $("#password").val('');
                $("#passwordInput").hide();
                $("#globalExportPopup").hide();  // 팝업 닫기
                $("#overlay").hide();  // 백그라운드 덮기 비활성화
            });

            // 암호화 옵션 체크박스 선택 시 비밀번호 입력 활성화
            $("#encryptCheckbox").change(function () {
                if (this.checked) {
                    $("#passwordInput").show();
                } else {
                    $("#passwordInput").hide();
                }
            });

            // 전체선택 기능
            $("#selectAll").click(function () {
                $("#objectSelectContainer input[type=checkbox]").prop("checked", true);
            });

            // Excel Export 실행
            $("#exportBtn").off("click").on("click", function () {
                var selectedObjects = [];
                $("#objectSelectContainer input[type=checkbox]:checked").each(function () {
                    selectedObjects.push($(this).val());
                });
                var encrypt = $("#encryptCheckbox").is(":checked");
                var password = $("#password").val();
                console.log(selectedObjects);
                if (selectedObjects.length > 0) {
                     exportSelectedObjects(selectedObjects, encrypt, password, layout.serverAddress);
                    //exportSelectedObjectss(selectedObjects, encrypt, password, layout.serverAddress);
                } else {
                    console.log("Export할 오브젝트를 선택해 주세요.");
                }
            });

            return qlik.Promise.resolve();
        }
    };

    // Export 가능한 오브젝트 필터링 및 드롭다운에 채우기
    async function populateExportableObjects() {
        var app = qlik.currApp();
        var exportableTypes = ["table", "pivot-table", "sn-pivot-table", "barchart", "piechart", "linechart", "combochart", "QuickTableViewer", "kpi"];

        var $container = $("#objectSelectTbody");
        $container.empty(); // 체크박스 목록 초기화

        // 현재 시트의 ID를 가져옴
        //var currentSheetId = qlik.navigation.getCurrentSheetId().sheetId;
        var currentSheetId = '';
        await qlik.currApp().getAppObjectList('sheet', function(reply) {
            reply.qAppObjectList.qItems.forEach(function(sheet) {
                var sheetObj = sheet;
                sheet.qData.cells.forEach(function(obj){
                    if(obj.name == extId){
                        currentSheetId = sheetObj.qInfo.qId; // 현재 확장 ID가 속한 시트 ID
                        return;
                    }
                });
                if(currentSheetId) return;
            });
            if(currentSheetId) return;
        });
        
        // objectMap이 비어있는 경우, 데이터를 새로 불러옴
        app.getObjectProperties(currentSheetId).then(function (props) {
                console.log(props);
                props.layout.qChildList.qItems.forEach(item => {
                    objectMap.set(item.qInfo.qId, item.qData.title || item.qInfo.qId);
                    if (exportableTypes.includes(item.qInfo.qType)) {
                        // $container.append(`
                        //       <div>
                        //           <input type="checkbox" value="${item.qInfo.qId}"> (${item.qInfo.qType})${item.qData.title ? item.qData.title + '_' : ''}${item.qInfo.qId}
                        //       </div>
                        //   `);
                        $container.append(`
                            <tr>
                                <td>  <input type="checkbox" value="${item.qInfo.qId}"> </td>
                                <td> ${item.qInfo.qType} </td>
                                <td> ${item.qData.title} </td>
                                <td> ${item.qInfo.qId} </td>
                            </tr>
                        `)
                    }
                });
            })
            .catch(function (error) {
                ////console.error("Error while fetching object properties:", error);
                console.log("오브젝트 데이터를 불러오는 데 문제가 발생했습니다.");
            });
    }

    // 안전한 fetch 요청 (에러 처리 포함)
    function safeFetch(url) {
        return fetch(url)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`); // 상태 코드가 오류인 경우 처리
                }
                return response.arrayBuffer();
            })
            .catch(error => {
                ////console.error("Fetch error:", error);
                console.log("파일을 가져오는 데 문제가 발생했습니다. 다시 시도해 주세요."); // 사용자에게 피드백
                throw error;  // 상위에서 처리할 수 있도록 오류를 다시 던짐
            });
    }
    
    function exportSelectedObjectss(objectIds, encrypt, password, severAddress){
        var exportPromises = objectIds.map(function (objId, index) {
            console.log(objId + ', ' + index);
        });
    }

    // Excel Export 기능 구현
    function exportSelectedObjects(objectIds, encrypt, password, serverAddress) {
        var app = qlik.currApp();
        var wb = XLSX.utils.book_new();

        var exportPromises = objectIds.map(function (objId, index) {
            return app.visualization.get(objId)
                .then(function (vis) {
                    return vis.exportData();
                })
                .then(function (result) {
                    return safeFetch(result);  // 안전한 fetch 요청
                })
                .then(function (data) {
                    // 파일 내용을 SheetJS로 파싱
                    var newWorkbook = XLSX.read(new Uint8Array(data), { type: "array" });
                    var sheetName = newWorkbook.SheetNames[0];  // 첫 번째 시트 이름 가져오기
                    var newSheet = newWorkbook.Sheets[sheetName];

                    // 새로운 시트를 통합 엑셀 파일에 추가
                    var filename = objectMap.get(objId) ? objectMap.get(objId) : objId;
                    filename = filename.substr(0, 31);
                    XLSX.utils.book_append_sheet(wb, newSheet, filename);
                })
                .catch(function (error) {
                    ////console.error("Error during object export:", error);
                    console.log("오브젝트 데이터를 내보내는 중 문제가 발생했습니다.");
                });
        });

        Promise.all(exportPromises)
            .then(function () {
                var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                var blob = new Blob([wbout], { type: 'application/octet-stream' });

                if (encrypt && serverAddress) {
                    var formData = new FormData();
                    formData.append("file", blob, "export.xlsx");
                    formData.append("password", password);

                    fetch(serverAddress + "/api/excel/protect", {
                        method: 'POST',
                        body: formData
                    })
                        .then(response => response.blob())
                        .then(encryptedBlob => {
                            var link = document.createElement('a');
                            link.href = window.URL.createObjectURL(encryptedBlob);
                            link.download = "Encrypted_Export.xlsx";
                            link.click();
                        })
                        .catch(error => {
                            ////console.error("Error during file encryption:", error);
                            console.log("파일 암호화 중 오류가 발생했습니다.");
                        });
                } else {
                    var link = document.createElement('a');
                    link.href = window.URL.createObjectURL(blob);
                    link.download = "Consolidated_Export.xlsx";
                    link.click();
                }
            })
            .catch(function (error) {
                ////console.error("Error during Excel export process:", error);
                console.log("엑셀 파일을 생성하는 중 문제가 발생했습니다.");
            });
    }
});

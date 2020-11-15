/* 
	3. 웹 페이지의 onload 이벤트에서 javascript 변수를 정의하며, Workbook 구조체를 호출하고 이전 단계에서 만든 div 요소의 DOM 개체에 제공하여 이 변수를 workbook 개체에 할당한다.
	또한 Workbook 개체는 인스턴스화하는 동안 설정할 수 있는 통합 문서의 다양한 속성을 지정하는 선택적 JSON 개체를 사용한다.
	또는 통합 문서가 만들어진 후 해당 옵션을 지정할 수 있다. 
*/
var spreadNS = GC.Spread.Sheets, spread;
window.onload = function() {
	// id가 "ss"인 DIV 요소에서 통합 문서 컨트롤 호스팅
	// GG.Spread.Sheets.Workbook 함수 사용
	spread = new spreadNS.Workbook(_getElementById('ss'));
	initSpread(spread);
	spread.suspendPaint();
	/*
	* findControl 정적 메서드를 사용하여 호스트 요소에서 스프레드 통합 문서 개체를 검색한다.
	* GC.Spread.Sheets.findControl 메서드를 사용하여 페이지의 통합 문서에 액세스하고 통합 문서를 호스팅하는 페이지의 DOM 요소를 제공할 수 있다.
    * var spread = GC.Spread.Sheets.findControl(_getElementById('ss'));
    */
	spread.resumePaint();
	
};

function initSpread(spread) {

	// 수식 텍스트 상자
	var fbx = new spreadNS.FormulaTextBox.FormulaTextBox(document.getElementById('formulaBar'));
	fbx.workbook(spread);
	
	// 하단 상태표시줄
    var statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(_getElementById('statusBar'));
    statusBar.bind(spread); // 바인딩
	//statusBar.dispose();  // 제거

	/******************************* 파일 - 가져오기 - 스프레드시트 파일(JSON) ******************************/

	/*  
		[참고 자료]
		https://demo.grapecity.co.kr/spreadjs/learn-spreadjs/features/workbook/json-serialization/purejs
		https://www.grapecity.com/spreadjs/docs/v13/online/json.html 
	*/

	// SSJSON 파일 가져오기
    _getElementById('loadSSJSON').addEventListener('click', function () {
		var target = _getElementById("fileSSJSON");	
		var file = target.files[0];
		
		// SSJSON 파일에서 텍스트 추출하기
		var reader = new FileReader();
		reader.onload = function(e) {
			console.log(e.target.result);
			
			spread.fromJSON(JSON.parse(e.target.result));	// 워크시트에 JSON 파일 뿌리기
		};
		reader.readAsText(file, 'utf8');
	});
	
	/************************************ 파일 - 가져오기 - Excel 파일 *************************************/

	/*  
		[참고자료]
		https://demo.grapecity.co.kr/spreadjs/learn-spreadjs/features/workbook/excel-import-export/purejs 
	*/

	// Excel IO 인스턴스 초기화
	var excelIo = new GC.Spread.Excel.IO(); 

	// 가져오기
    _getElementById('loadExcel').addEventListener('click', function () {
        var excelFile = _getElementById("fileDemo").files[0];
		var password = _getElementById('password').value;
		
        // Excel 파일 가져오기
        excelIo.open(excelFile, function (json) {   // open 메서드
            var workbookObj = json;
            spread.fromJSON(workbookObj);
        }, function (e) {
            // 에러 발생 시
            alert(e.errorMessage);
            if (e.errorCode === 2/*noPassword*/ || e.errorCode === 3 /*invalidPassword*/) {
                _getElementById('password').onselect = null;
            }
        }, {password: password});
	});

	/******************************* 파일 - 내보내기 - 스프레드시트 파일(JSON) ******************************/

	/* 
		[참고자료]
		https://stackoverflow.com/questions/19721439/download-json-object-as-a-file-from-browser 
	*/

	// 내보내기
    _getElementById('saveSSJSON').addEventListener('click', function () {
		var jsonStr = JSON.stringify(spread.toJSON(false));	// JSON 개체에 string 형식으로 저장

		_getElementById('download').setAttribute(
			"href", "data:text/json;charset=utf8," + encodeURIComponent(jsonStr)
		);
		_getElementById('download').click();
	});
	
	/*********************************** 파일 - 내보내기 - Excel 파일 ************************************/

	// 내보내기
    _getElementById('saveExcel').addEventListener('click', function () {

        /* var fileName = _getElementById('exportFileName').value; */
        var password = _getElementById('password').value;
        /* if (fileName.substr(-5, 5) !== '.xlsx') {
            fileName += '.xlsx';
        } */

        var json = spread.toJSON();

        // Excel 파일로 내보내기
        excelIo.save(json, function (blob) {    // save 메서드
            saveAs(blob, 'export.xlsx');
        }, function (e) {
            // 에러 발생 시
            console.log(e);
        }, {password: password});

	});
	
	/******************************* 파일 - 내보내기 - PDF 파일 ******************************/

	/*  
		[참고자료]
		https://www.grapecity.com/spreadjs/docs/v13/online/CustomPDFExport.html 
	*/

	// PDF 내보내기
    _getElementById('savePDF').addEventListener('click', function () {

		var sheet = spread.getActiveSheet();
		//sheet.setRowPageBreak(20,true);	// 행 나누기
		sheet.setColumnPageBreak(8,true);	// 열 나누기

        /* spread.savePDF(
            function (blob) {
                saveAs(blob, 'download.pdf');
            },
            console.log,
            {
                title: 'Test Title',
                author: 'Test Author',
                subject: 'Test Subject',
                keywords: 'Test Keywords',
                creator: 'test Creator'
			}); */

		spread.print();
    });
	
	/************************************ 홈 - 셀 - 형식 *************************************/

	// 행 숨기기
	_getElementById("chkRowHide").addEventListener('click', function () {
		var sheet = spread.getActiveSheet();
		var rowIndex = parseInt(_getElementById("rowIndex").value);

		if (!isNaN(rowIndex)) {
			sheet.setRowVisible(rowIndex, !this.checked);
		}
	});

	// 행 높이 자동 맞춤
	_getElementById("chkRowAutoFit").addEventListener('click', function () {
		var sheet = spread.getActiveSheet();
		var rowIndex = parseInt(_getElementById("rowIndex").value);

		if (!isNaN(rowIndex)) {
			var checked = this.checked;

			if (checked) {
				sheet.autoFitRow(rowIndex);
			}
		}
	});

	// 열 숨기기
	_getElementById("chkColumnHide").addEventListener('click', function () {
		var sheet = spread.getActiveSheet();
		var columnIndex = parseInt(_getElementById("columnIndex").value);

		if (!isNaN(columnIndex)) {
			sheet.setColumnVisible(columnIndex, !this.checked);
		}
	});
	
	// 열 너비 자동 맞춤
	_getElementById("chkColumnAutoFit").addEventListener('click', function () {
		var sheet = spread.getActiveSheet();
		var columnIndex = parseInt(_getElementById("columnIndex").value);

		if (!isNaN(columnIndex)) {
			var checked = this.checked;

			if (checked) {
				sheet.autoFitColumn(columnIndex);
			}
		}
	});

	/*********************************** 보기 - 표시/숨기기 ***********************************/

	// 행 머리글
	_getElementById('rowHeaderVisible_view').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.rowHeaderVisible = this.checked;	// 행 머리글
		_getElementById('rowHeaderVisible_setting').checked = this.checked;
		sheet.resumePaint();
	});

	// 열 머리글
	_getElementById('colHeaderVisible_view').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.colHeaderVisible = this.checked;	// 열 머리글
		_getElementById('colHeaderVisible_setting').checked = this.checked;
		sheet.resumePaint();
	});

	// 세로 눈금선
	_getElementById('VerticalGridline_view').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.gridline.showVerticalGridline = this.checked;
		_getElementById('VerticalGridline_setting').checked = this.checked;
		sheet.resumePaint();
	});

	// 가로 눈금선
	_getElementById('HorizontalGridline_view').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.gridline.showHorizontalGridline = this.checked;
		_getElementById('HorizontalGridline_setting').checked = this.checked;
		sheet.resumePaint();
	});

	// 연속 탭
	_getElementById('tabstrip_visible_view').addEventListener('click', function() {
		spread.options.tabStripVisible = this.checked;	// 탭 스트립의 표시 제어
		_getElementById('tabstrip_visible_setting').checked = this.checked;

		spread.invalidateLayout();
		spread.repaint();
	});

	// 새 탭
	_getElementById('newtab_show_view').addEventListener('click', function() {
		spread.options.newTabVisible = this.checked;	// + 원형 버튼 표시 제어
		_getElementById('newtab_show_setting').checked = this.checked;

		spread.invalidateLayout();
		spread.repaint();
	});


	/********************************* 설정 - 분배 설정 - 일반 ********************************/

	// 픽셀 스크롤 체크 박스
    _getElementById("scrollByPixel").addEventListener("change", function () {
        spread.options.scrollByPixel = scrollByPixel.checked;	// 픽셀 스크롤 사용 제어
	});
	
	var scrollPixel = _getElementById("scrollPixel");	// 사용자가 입력한 픽셀 스크롤 수
    _getElementById("setScrollPixel").addEventListener("click", function () {
        spread.options.scrollPixel = parseInt(scrollPixel.value);	// 픽셀 스크롤 수 설정
	});
	
	// 행 크기 조정 모드
    var rowResizeMode = _getElementById("rowResizeMode");
    rowResizeMode.addEventListener("change", function () {
        spread.options.rowResizeMode = parseInt(rowResizeMode.options[rowResizeMode.selectedIndex].value, 10);
    });

    // 열 크기 조정 모드
    var columnResizeMode = _getElementById("columnResizeMode");
    _getElementById("columnResizeMode").addEventListener("change", function () {
        spread.options.columnResizeMode = parseInt(columnResizeMode.options[columnResizeMode.selectedIndex].value, 10);
    });

	/****************************** 설정 - 분배 설정 - 스크롤 막대 *****************************/

	// 가로 스크롤 막대
	_getElementById('showHorizontalScrollbar').addEventListener('change', function() {
		spread.options.showHorizontalScrollbar = this.checked;	// 가로 스크롤 막대 표시 제어
	});

	// 세로 스크롤 막대
	_getElementById('showVerticalScrollbar').addEventListener('change', function() {
		spread.options.showVerticalScrollbar = this.checked;	// 세로 스크롤 막대 표시 제어
	});

	// 스크롤 막대 최대값 맞춤
	_getElementById('scrollbarMaxAlign').addEventListener('change', function() {
		spread.options.scrollbarMaxAlign = this.checked;	// 행 또는 열이 있는 영역까지 스크롤 제한
	});

	// 스크롤 막대 최대값 표시
	_getElementById('scrollbarShowMax').addEventListener('change', function() {
		spread.options.scrollbarShowMax = this.checked;		// 활성 시트의 전체 행 또는 열 수에 따라 컨테이너 크기 계산하여 스크롤바 표시 제어
	});

	// Mobile 스크롤 막대 사용
    _getElementById("mobileScrollbar").addEventListener("change", function () {
        spread.options.scrollbarAppearance = mobileScrollbar.checked ? GC.Spread.Sheets.ScrollbarAppearance.mobile : GC.Spread.Sheets.ScrollbarAppearance.skin;
    });

	/****************************** 설정 - 분배 설정 - 탭 스트립 *******************************/

	// 탭 스트립 표시
	_getElementById('tabstrip_visible_setting').addEventListener('click', function() {
		spread.options.tabStripVisible = this.checked;	// 탭 스트립의 표시 제어
		_getElementById('tabstrip_visible_view').checked = this.checked;

		spread.invalidateLayout();
		spread.repaint();
	});

	// 탭 스트립 편집 가능
	_getElementById('tab_editable').addEventListener('click', function() {
		spread.options.tabEditable = this.checked;			// 시트 이름 변경 제어
		//spread.options.allowSheetReorder = this.checked;	// 시트 순서 제어
	});

	// 새 탭 표시
	_getElementById('newtab_show_setting').addEventListener('click', function() {
		spread.options.newTabVisible = this.checked;	// + 원형 버튼 표시 제어
		_getElementById('newtab_show_view').checked = this.checked;

		spread.invalidateLayout();
		spread.repaint();
	});

	// 탭 스트립 비율
	_getElementById('setTabStripRatio').addEventListener('click', function() {
		var ratio = parseFloat(_getElementById('tabstrip_ratio').value);
		if (!isNaN(ratio)) {
			spread.options.tabStripRatio = ratio;	// 탭 스트립에 할당되는 가로 공간 크기를 지정하는 백분율 값(0.x)
		}
	});

	/***************************** 설정 - 시트 설정 - 일반 *******************************/

	// 틀고정 설정
	_getElementById('btnSetFrozenLine').addEventListener('click', function() {
		var sheet = spread.getActiveSheet();

		// 고정된 행 수
		if (_getElementById('rowCount').value) {
			var rowCount = parseInt(_getElementById('rowCount').value);
			sheet.frozenRowCount(rowCount);
		}
		// 후행 고정된 행 수
		if (_getElementById('trailingRowCount').value) {
			var trailingRowCount = parseInt(_getElementById('trailingRowCount').value);
			sheet.frozenTrailingRowCount(trailingRowCount);
		}
		// 고정된 열 수
		if (_getElementById('columnCount').value) {
			var columnCount = parseInt(_getElementById('columnCount').value);
			sheet.frozenColumnCount(columnCount);
		}
		// 후행 고정된 열 수
		if (_getElementById('trailingColumnCount').value) {
			var trailingColumnCount = parseInt(_getElementById('trailingColumnCount').value);
			sheet.frozenTrailingColumnCount(trailingColumnCount);
		}
	});

	/***************************** 설정 - 시트 설정 - 눈금선 *******************************/

	// 가로 눈금선
	_getElementById('HorizontalGridline_setting').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.gridline.showHorizontalGridline = this.checked;
		_getElementById('HorizontalGridline_view').checked = this.checked;
		sheet.resumePaint();
	});

	// 세로 눈금선
	_getElementById('VerticalGridline_setting').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.gridline.showVerticalGridline = this.checked;
		_getElementById('VerticalGridline_view').checked = this.checked;
		sheet.resumePaint();
	});

	/***************************** 설정 - 시트 설정 - 머리글 *******************************/

	// 열 머리글
	_getElementById('colHeaderVisible_setting').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.colHeaderVisible = this.checked;	// 열 머리글
		_getElementById('colHeaderVisible_view').checked = this.checked;
		sheet.resumePaint();
	});

	// 행 머리글
	_getElementById('rowHeaderVisible_setting').addEventListener('change', function() {
		var sheet = spread.getActiveSheet();
		sheet.suspendPaint();
		sheet.options.rowHeaderVisible = this.checked;	// 행 머리글
		_getElementById('rowHeaderVisible_view').checked = this.checked;
		sheet.resumePaint();
	});

	// 자동 텍스트 설정
	_getElementById("btnSetAutoText").addEventListener('click',function () {
		var sheet = spread.getActiveSheet();
		var headerType = document.querySelector("input[name='headerType']:checked").value,
			obj=_getElementById("headerAutoTextType"),
			headerAutoTextType = obj.options[obj.selectedIndex].value;

		if (headerAutoTextType) {
			headerAutoTextType = spreadNS.HeaderAutoText[headerAutoTextType];

			if (!(headerAutoTextType === undefined)) {
				switch (headerType) {
					case "row":
						sheet.options.rowHeaderAutoText = headerAutoTextType;
						break;
					case "column":
						sheet.options.colHeaderAutoText = headerAutoTextType;
						break;
				}
			}
		}
	});

	/***************************** 설정 - 시트 설정 - 시트 탭 *******************************/

	// 시트 탭 색
	_getElementById('setSheetTabColor').addEventListener('click', function() {
		var sheet = spread.getActiveSheet();
		if (sheet) {
			var color = _getElementById('sheetTabColor').value;
			sheet.options.sheetTabColor = color;	// 시트 탭의 색상 변경
		}
	});

	/*************************************** 검색 ****************************************/

	// 다음 찾기
	_getElementById('btnFindNext').onclick = function () {
        var sheet = spread.getActiveSheet();						// 현재 활성화된 시트
        var searchCondition = getSearchCondition();					// 검색 조건
        var within = _getElementById('searchWithin').value;	// 범위
        var searchResult = null;
        if (within == "sheet") {
			// 검색 범위가 시트일 때
            var sels = sheet.getSelections();
            if (sels.length > 1) {
                searchCondition.searchFlags |= spreadNS.Search.SearchFlags.blockRange;
            } else if (sels.length == 1) {
                var spanInfo = getSpanInfo(sheet, sels[0].row, sels[0].col);
                if (sels[0].rowCount != spanInfo.rowSpan && sels[0].colCount != spanInfo.colSpan) {
                    searchCondition.searchFlags |= spreadNS.Search.SearchFlags.blockRange;
                }
            }
            searchResult = getResultSearchinSheetEnd(searchCondition);
            if (searchResult == null || searchResult.searchFoundFlag == spreadNS.Search.SearchFoundFlags.none) {
                searchResult = getResultSearchinSheetBefore(searchCondition);
            }
        } else if (within == "workbook") {
			// 검색 범위가 통합 문서일 때
            searchResult = getResultSearchinSheetEnd(searchCondition);
            if (searchResult == null || searchResult.searchFoundFlag == spreadNS.Search.SearchFoundFlags.none) {
                searchResult = getResultSearchinWorkbookEnd(searchCondition);
            }
            if (searchResult == null || searchResult.searchFoundFlag == spreadNS.Search.SearchFoundFlags.none) {
                searchResult = getResultSearchinWorkbookBefore(searchCondition);
            }
            if (searchResult == null || searchResult.searchFoundFlag == spreadNS.Search.SearchFoundFlags.none) {
                searchResult = getResultSearchinSheetBefore(searchCondition);
            }
        }

        if (searchResult != null && searchResult.searchFoundFlag != spreadNS.Search.SearchFoundFlags.none) {
			// 찾았을 때
            spread.setActiveSheetIndex(searchResult.foundSheetIndex);
            var sheet = spread.getActiveSheet();
            sheet.setActiveCell(searchResult.foundRowIndex, searchResult.foundColumnIndex);
            if ((searchCondition.searchFlags & spreadNS.Search.SearchFlags.blockRange) == 0) {
                sheet.setActiveCell(searchResult.foundRowIndex, searchResult.foundColumnIndex, 1, 1);
            }
            //scrolling
            if (searchResult.foundRowIndex < sheet.getViewportTopRow(1)
                || searchResult.foundRowIndex > sheet.getViewportBottomRow(1)
                || searchResult.foundColumnIndex < sheet.getViewportLeftColumn(1)
                || searchResult.foundColumnIndex > sheet.getViewportRightColumn(1)
            ) {
                sheet.showCell(searchResult.foundRowIndex,
                    searchResult.foundColumnIndex,
                    spreadNS.VerticalPosition.center,
                    spreadNS.HorizontalPosition.center);
            } else {
                sheet.repaint();
            }
        } else {
            // 찾지 못했을 때
            alert('검색하는 항목을 찾지 못했습니다.');
        }
    };

	// 데이터 바인딩
	/* var sd = dataSource;
	var sheet = spread.getActiveSheet();
	if (sd.length > 0) {
		sheet.setDataSource(sd);
	} */
}

function _getElementById(id) {
	return document.getElementById(id);
}

function getSpanInfo(sheet, row, col) {
    var span = sheet.getSpans(new spreadNS.Range(row, col, 1, 1));
    if (span.length > 0) {
        return {rowSpan: span[0].rowCount, colSpan: span[0].colCount};
    } else {
        return {rowSpan: 1, colSpan: 1};
    }
}

function getResultSearchinSheetEnd(searchCondition) {
    var sheet = spread.getActiveSheet();
    searchCondition.startSheetIndex = spread.getActiveSheetIndex();
    searchCondition.endSheetIndex = spread.getActiveSheetIndex();

    if (searchCondition.searchOrder == spreadNS.Search.SearchOrder.zOrder) {
        searchCondition.findBeginRow = sheet.getActiveRowIndex();
        searchCondition.findBeginColumn = sheet.getActiveColumnIndex() + 1;
    } else if (searchCondition.searchOrder == spreadNS.Search.SearchOrder.nOrder) {
        searchCondition.findBeginRow = sheet.getActiveRowIndex() + 1;
        searchCondition.findBeginColumn = sheet.getActiveColumnIndex();
    }

    if ((searchCondition.searchFlags & spreadNS.Search.SearchFlags.blockRange) > 0) {
        var sel = sheet.getSelections()[0];
        searchCondition.rowStart = sel.row;
        searchCondition.columnStart = sel.col;
        searchCondition.rowEnd = sel.row + sel.rowCount - 1;
        searchCondition.columnEnd = sel.col + sel.colCount - 1;
    }
    var searchResult = spread.search(searchCondition);
    return searchResult;
}

function getResultSearchinSheetBefore(searchCondition) {
    var sheet = spread.getActiveSheet();
    searchCondition.startSheetIndex = spread.getActiveSheetIndex();
    searchCondition.endSheetIndex = spread.getActiveSheetIndex();
    if ((searchCondition.searchFlags & spreadNS.Search.SearchFlags.blockRange) > 0) {
        var sel = sheet.getSelections()[0];
        searchCondition.rowStart = sel.row;
        searchCondition.columnStart = sel.col;
        searchCondition.findBeginRow = sel.row;
        searchCondition.findBeginColumn = sel.col;
        searchCondition.rowEnd = sel.row + sel.rowCount - 1;
        searchCondition.columnEnd = sel.col + sel.colCount - 1;
    } else {
        searchCondition.rowStart = -1;
        searchCondition.columnStart = -1;
        searchCondition.findBeginRow = -1;
        searchCondition.findBeginColumn = -1;
        searchCondition.rowEnd = sheet.getActiveRowIndex();
        searchCondition.columnEnd = sheet.getActiveColumnIndex();
    }

    var searchResult = spread.search(searchCondition);
    return searchResult;
}

function getResultSearchinWorkbookEnd(searchCondition) {
    searchCondition.rowStart = -1;
    searchCondition.columnStart = -1;
    searchCondition.findBeginRow = -1;
    searchCondition.findBeginColumn = -1;
    searchCondition.rowEnd = -1;
    searchCondition.columnEnd = -1;
    searchCondition.startSheetIndex = spread.getActiveSheetIndex() + 1;
    searchCondition.endSheetIndex = -1;
    var searchResult = spread.search(searchCondition);
    return searchResult;
}

function getResultSearchinWorkbookBefore(searchCondition) {
    searchCondition.rowStart = -1;
    searchCondition.columnStart = -1;
    searchCondition.findBeginRow = -1;
    searchCondition.findBeginColumn = -1;
    searchCondition.rowEnd = -1;
    searchCondition.columnEnd = -1;
    searchCondition.startSheetIndex = -1
    searchCondition.endSheetIndex = spread.getActiveSheetIndex() - 1;
    var searchResult = spread.search(searchCondition);
    return searchResult;
}

// 검색 조건 가져오는 함수
function getSearchCondition() {
    var searchCondition = new spreadNS.Search.SearchCondition();					// 검색 조건
    var findWhat = _getElementById('txtSearchWhat').value;					// 찾는 내용
    var within = _getElementById('searchWithin').value;						// 범위(시트/통합 문서)
    var order = _getElementById('searchOrder').value;						// 검색(행/열)
    var lookin = _getElementById('searchLookin').value;						// 찾는 위치(값/수식)
    var matchCase = _getElementById('chkSearchMachCase').checked;			// 대/소문자 일치
    var matchEntire = _getElementById('chkSearchMachEntire').checked;		// 정확히 일치
    var useWildCards = _getElementById('chkSearchUseWildCards').checked;	// 와일드카드 사용

    searchCondition.searchString = findWhat;	// 검색 문자 = 찾는 내용
    if (within == "sheet") {
		// 범위가 시트일 때
        searchCondition.startSheetIndex = spread.getActiveSheetIndex();
        searchCondition.endSheetIndex = spread.getActiveSheetIndex();
    }
    if (order == "norder") {
		// 열에서 검색할 때
        searchCondition.searchOrder = spreadNS.Search.SearchOrder.nOrder;
    } else {
		// 행에서 검색할 때
        searchCondition.searchOrder = spreadNS.Search.SearchOrder.zOrder;
    }
    if (lookin == "formula") {
		// 찾는 위치가 수식일 때
        searchCondition.searchTarget = spreadNS.Search.SearchFoundFlags.cellFormula;
    } else {
		// 찾는 위치가 값일 때
        searchCondition.searchTarget = spreadNS.Search.SearchFoundFlags.cellText;
    }

    if (!matchCase) {
		// 대/소문자가 일치하지 않아도 될 때
        searchCondition.searchFlags |= spreadNS.Search.SearchFlags.ignoreCase;
    }
    if (matchEntire) {
		// 정확히 일치해야할 때
        searchCondition.searchFlags |= spreadNS.Search.SearchFlags.exactMatch;
    }
    if (useWildCards) {
		// 와일드카드 사용할 때
        searchCondition.searchFlags |= spreadNS.Search.SearchFlags.useWildCards;
    }

    return searchCondition;	// 검색 조건 반환
}
/**
 * 제 품: IBSheet8 - Common Plugin
 * 버 전: v0.0.0 (20200108-1828)
 * 회 사: (주)아이비리더스
 * 주 소: https://www.ibsheet.com
 * 전 화: (02) 2621-2288~9
 */
(function(window, document) {
/*CommonOptions 설정
 * 모든 시트에 동일하게 적용하고자 하는 설정을 CommonOptions에 등록합니다.
 * 해당 파일은 반드시 ibsheet.js 파일보다 뒤에 include 되어야 합니다.
 */
var _IBSheet = window['IBSheet'];
if (_IBSheet == null) {
  throw new Error('[ibsheet-common] undefined global object: IBSheet');
}

// IBSheet Plugins Object
var Fn = _IBSheet['Plugins'];

if (!Fn.PluginVers) Fn.PluginVers = {};
Fn.PluginVers.ibcommon = {
  name: 'ibcommon',
  version: '0.0.0-20200108-1828'
};

_IBSheet.CommonOptions = {
  Cfg: {
    Export: {
      Url: '../assets/ibsheet/jsp/'
    }, // 엑셀다운 URL
    Alternate: 2, // 홀짝 행에 대한 배경색 설정
    InfoRowConfig: {
      Visible: 1,
      Layout: ['Count'],
      Space: 'Top'
    }, // 건수 정보 표시
    GroupFormat: ' <span style=\'color:red\'>{%s}</span> <span style=\'color:blue\'>({%c}건)</span>', // 그룹핑 컬럼명은 빨강색, 건수는 파란색으로 표시

    HeaderMerge: 1, // 헤더영역 자동 병합
    PrevColumnMerge: 3, // 앞컬럼 기준 병합 사용 여부

    SearchCells: 1, // 찾기 기능 셀단위/행단위 선택

    MaxPages: 6, // SearchMode:2인 경우 한번에 갖고 있는 페이지 수(클수록 브라우져의 부담이 커짐)
    MaxSort: 3, // 최대 소팅 가능 컬럼수(4개 이상인 경우 느려질 수 있음)

    StorageSession: 1, // 개인화 기능(컬럼정보 저장) 사용 여부
    StorageKeyPrefix: window['sampleName'] ? window['sampleName'] : location.href // 저장 키 prefix 설정
  },
  Def: {
    Header: { //헤더 영역 행에 대한 설정
      Menu: {
        Items: [{
            'Name': '컬럼 감추기'
          },
          {
            'Name': '컬럼 감추기 취소'
          },
          {
            'Name': '컬럼 정보 저장'
          },
          {
            'Name': '컬럼 정보 저장 취소'
          }
        ],
        'OnSave': function (item, data) {
          switch (item.Name) {
            case '컬럼 감추기':
              var col = item.Owner.Col;
              this.Sheet.hideCol(col, 1);
              break;
            case '컬럼 감추기 취소':
              this.Sheet.showCol();
              break;
            case '컬럼 정보 저장':
              this.Sheet.saveCurrentInfo();
              break;
            case '컬럼 정보 저장 취소':
              this.Sheet.clearCurrentInfo();
              this.Sheet.showMessageTime({
                message: '컬럼 정보를 삭제하였습니다.<br>새로고침하시면 초기 설정의 시트를 확인하실 수 있습니다.'
              });
              break;
          }
        }
      }
    },

    Row: { //데이터 영역 모든 행에 대한 설정
      // AlternateColor:"#F1F1F1",  //짝수행에 대한 배경색
      // Menu:{ //마우스 우측버특 클릭시 보여지는 메뉴 설정 (메뉴얼에서 Appedix/Menu 참고)
      //   "Items":[
      //     {"Name":"다운로드","Caption":1},
      //     {"Name":"Excel","Value":"xls"},
      //     {"Name":"text","Value":"txt"},
      //     {"Name":"pdf","Value":"pdf"},
      //     // {"Name":"-"},
      //     {"Name":"데이터 수정","Caption":1},
      //     {"Name":"데이터 추가/제거",Menu:1,"Items":[
      //       {"Name":"위로 행 추가","Value":"addAbove"},
      //       {"Name":"아래로 행 추가","Value":"addBelow"},
      //       {"Name":"행 삭제","Value":"del"}
      //     ]},
      //     {"Name":"데이터 이동",Menu:1,"Items":[
      //       {"Name":"위로로 이동","Value":"moveAbove"},
      //       {"Name":"아래로 이동","Value":"moveBelow"},
      //     ]}

      //   ],
      //   "OnSave":function(item,data){//메뉴 선택시 발생 이벤트
      //     switch(item.Value){
      //       case 'xls':
      //         try{
      //           this.Sheet.down2Excel({FileName:"test.xlsx",SheetDesign:1});
      //         }catch(e){
      //           if(e.message.indexOf("down2Excel is not a function")>-1){
      //               console.log("%c 경고","color:#FF0000"," : ibsheet-excel.js 파일이 필요합니다.");
      //           }
      //         }
      //         break;
      //       case 'txt':
      //         try{
      //           this.Sheet.down2Text();
      //         }catch(e){
      //           if(e.message.indexOf("down2Text is not a function")>-1){
      //             console.log("%c 경고","color:#FF0000"," : ibsheet-excel.js 파일이 필요합니다.");
      //           }
      //         }
      //         break;
      //       case 'pdf':
      //         try{
      //           this.Sheet.down2Pdf();
      //         }catch(e){
      //           if(e.message.indexOf("down2Pdf is not a function")>-1){
      //             console.log("%c 경고","color:#FF0000"," : ibsheet-excel.js 파일이 필요합니다.");
      //           }
      //         }
      //         break;
      //       case 'addAbove'://위로 추가
      //         var nrow = item.Owner.Row;
      //         this.Sheet.addRow({next:nrow});
      //         break;
      //       case 'addBelow'://아래추가
      //         var nrow = this.Sheet.getNextRow(item.Owner.Row);
      //         this.Sheet.addRow({next:nrow});
      //         break;
      //       case 'del'://삭제
      //         var row = item.Owner.Row;
      //         this.Sheet.deleteRow(row);
      //         break;

      //       case 'moveAbove'://위로 이동
      //           var row = item.Owner.Row;
      //           var nrow = this.Sheet.getPrevRow(item.Owner.Row);
      //           this.Sheet.moveRow({row:row,next:nrow});
      //         break;
      //       case 'moveBelow'://아래로 이동
      //           var row = item.Owner.Row;
      //           var nrow = this.Sheet.getNextRow(this.Sheet.getNextRow(item.Owner.Row));
      //           this.Sheet.moveRow({row:row,next:nrow});
      //         break;
      //     }
      //   }
      // }
    }
  },
  Events: {
    'onKeyDown': function (evtParam) {
      // Ctrl+Shift+F 입력시 찾기 창 오픈
      if (evtParam.prefix == 'ShiftCtrl' && evtParam.key == 70) {
        evtParam.sheet.showFindDialog();
      } else if (evtParam.prefix == 'CtrlAlt' && evtParam.key == 84) {
        evtParam.sheet.createPivotDialog();
      }
    },
  }
};

window.IB_Preset = {
  // 날짜 시간 포멧
  'YMD': {
    Type: 'Date',
    Align: 'Center',
    Width: 110,
    Format: 'yyyy/MM/dd',
    DataFormat: 'yyyyMMdd',
    EditFormat: 'yyyyMMdd',
    Size: 8,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>년,월,일 순으로 숫자만 입력해 주세요.</span>'
  },
  'YM': {
    Type: 'Date',
    Align: 'Center',
    Width: 80,
    Format: 'yyyy/MM',
    DataFormat: 'yyyyMM',
    EditFormat: 'yyyyMM',
    Size: 6,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>년,월 순으로 숫자만 입력해 주세요.</span>'
  },
  'MD': {
    Type: 'Date',
    Align: 'Center',
    Width: 60,
    Format: 'MM/dd',
    EditFormat: 'MMdd',
    DataFormat: 'MMdd',
    Size: 4,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>월,일 순으로 숫자만 입력해 주세요.</span>'
  },
  'HMS': {
    Type: 'Date',
    Align: 'Center',
    Width: 70,
    Format: 'HH:mm:ss',
    EditFormat: 'HHmmss',
    DataFormat: 'HHmmss',
    Size: 8,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>시,분,초 순으로 8개 숫자만 입력해 주세요.</span>'
  },
  'HM': {
    Type: 'Date',
    Align: 'Center',
    Width: 70,
    Format: 'HH:mm',
    EditFormat: 'HHmm',
    DataFormat: 'HHmm',
    Size: 6,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>시,분 순으로 4개 숫자만 입력해 주세요.</span>'
  },
  'YMDHMS': {
    Type: 'Date',
    Align: 'Center',
    Format: 'yyyy/MM/dd HH:mm:ss',
    Width: 150,
    EditFormat: 'yyyyMMddHHmmss',
    DataFormat: 'yyyyMMddHHmmss',
    Size: 14,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>숫자만 입력(ex:20190514153020)</span>'
  },
  'YMDHM': {
    Type: 'Date',
    Align: 'Center',
    Format: 'yyyy/MM/dd HH:mm',
    Width: 150,
    EditFormat: 'yyyyMMddHHmm',
    DataFormat: 'yyyyMMddHHmm',
    Size: 12,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>숫자만 입력(ex:201905141530)</span>'
  },
  'MDY': {
    Type: 'Date',
    Align: 'Center',
    Format: 'MM-dd-yyyy',
    Width: 110,
    EditFormat: 'MMddyyyy',
    DataFormat: 'yyyyMMdd',
    Size: 8,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>월,일,년 순으로 숫자만 입력해 주세요.</span>'
  },
  'DMY': {
    Type: 'Date',
    Align: 'Center',
    Format: 'dd-MM-yyyy',
    Width: 110,
    EditFormat: 'ddMMyyyy',
    DataFormat: 'yyyyMMdd',
    Size: 8,
    EditMask: '^\\d*$',
    EmptyValue: '<span style=\'color:#AAA\'>일,월,년 순으로 숫자만 입력해 주세요.</span>'
  },

  // 숫자 포멧
  'Integer': {
    Type: 'Int',
    Align: 'Right',
    Format: '#,##0',
    Width: 100
  },
  'NullInteger': {
    Type: 'Int',
    Align: 'Right',
    Format: '#,###',
    Width: 100
  },
  'Float': {
    Type: 'Float',
    Align: 'Right',
    Format: '#,##0.######',
    Width: 100
  },
  'NullFloat': {
    Type: 'Float',
    Align: 'Right',
    Format: '#,###.######',
    Width: 100
  },

  // 기타포멧
  'IdNo': {
    Type: 'Text',
    Align: 'Center',
    CustomFormat: 'IdNo'
  },
  'SaupNo': {},
  'PostNo': {},
  'CardNo': {},
  'PhoneNo': {},
  'Number': {},

  // ibsheet7 migration
  // Status Type
  'STATUS': {
    Type: 'Text',
    Align: 'Center',
    Width: 50,
    Formula: 'Row.Deleted ? \'D\' : Row.Added ? \'I\' : Row.Changed ? \'U\' : \'R\'',
    Format: {
      'I': '입력',
      'U': '수정',
      'D': '삭제',
      'R': ''
    }
  },
  // DelCheck Type
  'DelCheck': {
    Type: 'Bool',
    Width: 50,
    OnClick: function(evtParam){
    	//부모가 체크되어 있는 경우 더 이상 진행하지 않는다.
    	var chked = !(evtParam.row[evtParam.col]);
    	var prows = evtParam.sheet.getParentRows( evtParam.row);
    	if(!chked && prows[0] && prows[0][evtParam.col]) return true;	
    },
    OnChange: function (evtParam) {
    	var chked = evtParam.row[evtParam.col];
    	//신규행에 대해서는 즉시 삭제한다.
      if (evtParam.row.Added) {
        setTimeout(function () {
          evtParam.sheet.removeRow(evtParam.row);
        }, 30);
      } else {
      	//행을 삭제 상태로 변경
        evtParam.sheet.deleteRow(evtParam.row, evtParam.row[evtParam.col]);
        //자식행 추출
        var rows = evtParam.sheet.getChildRows(evtParam.row);
        rows.push(evtParam.row);
        
        //모두 체크하고 편집 불가로 변경
        for(var i=0;i<rows.length;i++){
        	var row = rows[i];
        	evtParam.sheet.setValue (row ,evtParam.col, chked, 0 );
         	row.CanEdit = !evtParam.row[evtParam.col];
         	if (!row[evtParam.col+'CanEdit']) {
		        row[evtParam.col+'CanEdit'] = true;
		      }
         	evtParam.sheet.refreshRow(row);	
        }
      }
    }
  }
};

function clone(obj) {
  if (obj === null || typeof (obj) !== 'object') return obj;
  var copy = obj.constructor();
  for (var attr in obj) {
    if (obj.hasOwnProperty(attr)) {
      copy[attr] = clone(obj[attr]);
    }
  }
  return copy;
}

/*
ibsheet7 migration functions
*/
if (!_IBSheet.v7) _IBSheet.v7 = {};

/*
 * ibsheet7 AcceptKey 속성 대응
 * param list
 * objColumn : 시트 생성시 Cols객체의 컬럼
 * str : ibsheet7 AcceptKeys에 정의했던 스트링
 */
_IBSheet.v7.convertAcceptKeys = function (objColumn, str) {
  // EditMask를 통해 AcceptKeys를 유사하게 구현
  var acceptKeyArr = str.split('|');
  var mask = '';

  for (var i = 0; i < acceptKeyArr.length; i++) {
    switch (acceptKeyArr[i]) {
      case 'E':
        mask += '|\\w';
        break;
      case 'N':
        mask += '|\\d';
        break;
      case 'K':
        mask += '|\\u3131-\\u314e|\\u314f-\\u3163|\\uac00-\\ud7a3';
        break;
      default:
        if (acceptKeyArr[i].substring(0, 1) == '[' && acceptKeyArr[i].substring(acceptKeyArr[i].length - 1) == ']') {
          var otherKeys = acceptKeyArr[i].substring(1, acceptKeyArr[i].length - 1);
          for (var x = 0; x < otherKeys.length; x++) {
            if (otherKeys[x] == '.' || otherKeys[x] == '-') {
              mask += '|\\' + otherKeys[x];
            } else {
              mask += '|' + otherKeys[x];
            }
          }
        }
        break;
    }
  }
  objColumn.EditMask = '^[' + mask.substring(1) + ']*$';
};

//Date Format migration
//exam)
/*
//데이터 로드 이벤트에서 호출합니다.
options.Events.onBeforeDataLoad:function(obj){
  //날짜포맷 컬럼의 값을 ibsheet8에 맞게 변경하여 로드시킴
  IBSheet.v7.convertDateFormat(obj);
}
*/
_IBSheet.v7.convertDateFormat = function (obj) {
  var cdata = obj.data;
  var changeCol = {};
  //날짜 컬럼에 대한 포맷을 별도로 저장
  var cols = obj.sheet.getCols();
  for (var i = 0; i < cols.length; i++) {
    var colName = cols[i];

    if (obj.sheet.Cols[colName].Type == 'Date') {
      //DataFormat이 없으면 EditFormat 이나 포맷에서 알파벳만 추출
      var format = (obj.sheet.Cols[colName].DataFormat) ? obj.sheet.Cols[colName].DataFormat : (obj.sheet.Cols[colName].EditFormat) ? obj.sheet.Cols[colName].EditFormat : obj.sheet.Cols[colName].Format.replace(/([^A-Za-z])+/g, '');
      changeCol[colName] = {
        format: format,
        length: format.length
      };
    }
  }

  if (Object.keys(changeCol).length !== 0) {
    var changeColKeys = Object.keys(changeCol);

    //DataFormat의 길이만큼 문자열을 자름
    for (var row = 0; row < cdata.length; row++) {
      for (var colName in cdata[row]) {
        if (changeColKeys.indexOf(colName) > -1) {
          // 문자열만 처리
          if (typeof ((cdata[row])[colName]) == 'string') {
            //실제 값
            var v = (cdata[row])[colName];
            //MMdd만 값이 8자리 이상이면 중간에 4자리만 pick
            if (changeCol[colName].format == 'MMdd' && v.length != 4) {
              if (v.length > 7) {
                v = v.substr(4, 4);
              }
            } else {
              //일반적으로 모두 포맷의 문자열 길이만큼 자름
              v = v.substr(0, changeCol[colName].length);
            }
            //수정한 값을 원래 위치에 대입
            (cdata[row])[colName] = v;
          }
        }
      }
    }
  }
};

/* ibsheet7의 Tree 구조 Json 데이터를 ibsheet8 형식에 맞게 파싱해주는 메소드 */
_IBSheet.v7.convertTreeData = function (data7) {
  if (this.MainCol) {
    var targetArr;
    var toString = Object.prototype.toString;
    var startLevel = 0;
    switch (toString.call(data7)) {
      case '[object Object]':
        if (!(data7['data'] || data7['Data']) ||
          toString.call((data7['data'] || data7['Data'])) !== '[object Array]')
          return false;
        targetArr = (data7['data'] || data7['Data']);
        break;
      case '[object Array]':
        targetArr = data7;
        break;
      default:
        return false;
    }

    targetArr = targetArr.reduce(function (accum, currentVal, curretIndex, array) {
      var cloneObj = clone(currentVal);
      if (cloneObj['HaveChild']) {
        cloneObj['Count'] = true;
        delete cloneObj['HaveChild'];
      }
      if (accum.length == 0) {
        startLevel = parseInt(cloneObj['Level']);
        delete cloneObj['Level'];
        accum.push(cloneObj);
      } else if (currentVal['Level'] <= startLevel) {
        startLevel = parseInt(cloneObj['Level']);
        delete cloneObj['Level'];
        accum.push(cloneObj);
      } else if (currentVal['Level']) {
        var parent = accum[accum.length - 1];
        for (var i = startLevel; i < parseInt(currentVal['Level']); i++) {
          if (i === parseInt(currentVal['Level']) - 1) {
            if (!parent.Items) {
              parent.Items = [];
            }
            delete cloneObj['Level'];
            parent.Items.push(cloneObj);
          } else {
            parent = parent.Items[parent.Items.length - 1];
          }
        }
      }
      return accum;
    }, []);

    delete data7['Data'];
    data7['data'] = targetArr;
  }

  return data7;
};
}(window, document));

var wb = XLSX.utils.book_new();

const bank_account_code = {
  '국민': "004",
  '경남': "039",
  '광주': "034",
  '기업': "003",
  '농협': "011",
  '대구': "031",
  '부산': "032",
  '산업': "002",
  '저축': "050",
  '새마을금고': "045",
  '수협': "007",
  '신한': "088",
  '신협': "048",
  "씨티": '027',
  "하나": '081',
  "우리": '020',
  "우체국": '071',
  "전북": '037',
  "제일": '023',
  "제주": '035',
  "HSBC": '054',
  "케이뱅크": "089",
  "카카오": "090"
};

function downloadAsXLSX(jsonData) {
  let selectedData = jsonData.map(item => {
    let accountParts = item.bank_account_number.split(' ');
    let bankName = accountParts.shift();
    let accountCode = bank_account_code[bankName];

    return {
      계좌코드: accountCode,
      계좌번호: accountParts.join(' '),
      지급금액: item.real_amount.toLocaleString(),
      정산일: `케이저${item.month}월`,
      저작권자: item.holder
    };
  });

  // 워크북 생성
  const wb = XLSX.utils.book_new();

  // 워크시트 생성
  const ws = XLSX.utils.json_to_sheet(selectedData, { 
    header: ['계좌코드', '계좌번호', '지급금액', '정산일', '저작권자'], 
    skipHeader: true 
  });

  // 워크북에 워크시트 추가
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  // 파일 이름 동적으로 생성
  const fileName = `케이저${jsonData[0].month}월.xlsx`

  // 파일 저장 (웹에서 자동으로 다운로드)
  XLSX.writeFile(wb, fileName);
}

// 버튼 클릭 이벤트 처리기 설정
document.addEventListener('DOMContentLoaded', function () {
  const downloadButton = document.getElementById('download-button');
  downloadButton.addEventListener('click', function () {
      fetch('https://jigpu.com/v0.01/?_action=statistics_summary_by_holder&_action_type=load&_plugin=keiser')
          .then(response => response.json())
          .then(data => {
              downloadAsXLSX(data.list);
          })
          .catch(error => {
              console.error('Error:', error);
          });
  });
});


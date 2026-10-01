/**
 * ATELIER SEOUL - 온라인 쇼핑몰 사용자 정보 입력 및 엑셀 저장
 * 단독 JavaScript 파일 (script.js)
 */

document.addEventListener('DOMContentLoaded', () => {
  // 1. DOM 요소 취득
  const form = document.getElementById('orderForm');
  const btnSubmit = document.getElementById('btnSubmit');
  const btnExportExcelDirect = document.getElementById('btnExportExcelDirect');
  const btnFillDemo = document.getElementById('btnFillDemo');
  const btnReset = document.getElementById('btnReset');

  // 사용자 정보 입력 필드
  const nameInput = document.getElementById('userName');
  const phoneInput = document.getElementById('userPhone');
  const emailInput = document.getElementById('userEmail');
  const domainSelect = document.getElementById('emailDomainSelect');

  // 배송지 정보 필드
  const sameAsRecipientCheck = document.getElementById('sameAsRecipient');
  const recipientGroup = document.getElementById('recipientGroup');
  const recipientNameInput = document.getElementById('recipientName');
  const recipientPhoneInput = document.getElementById('recipientPhone');

  const zipcodeInput = document.getElementById('zipcode');
  const roadAddressInput = document.getElementById('roadAddress');
  const detailAddressInput = document.getElementById('detailAddress');
  const btnSearchAddress = document.getElementById('btnSearchAddress');

  const deliverySelect = document.getElementById('deliveryRequest');
  const customDeliveryInput = document.getElementById('customDeliveryRequest');

  // 약관 동의 체크박스
  const agreePrivacyCheck = document.getElementById('agreePrivacy');
  const agreeOrderCheck = document.getElementById('agreeOrder');

  // 에러 메시지 요소들
  const nameError = document.getElementById('nameError');
  const phoneError = document.getElementById('phoneError');
  const emailError = document.getElementById('emailError');
  const addressError = document.getElementById('addressError');
  const recipientNameError = document.getElementById('recipientNameError');
  const recipientPhoneError = document.getElementById('recipientPhoneError');
  const privacyError = document.getElementById('privacyError');
  const orderError = document.getElementById('orderError');

  // 모달 요소들
  const addressModal = document.getElementById('addressModal');
  const btnCloseAddressModal = document.getElementById('btnCloseAddressModal');
  const addressSearchInput = document.getElementById('addressSearchInput');
  const addressResultList = document.getElementById('addressResultList');

  const privacyModal = document.getElementById('privacyModal');
  const btnOpenPrivacy = document.getElementById('btnOpenPrivacy');
  const btnClosePrivacyModal = document.getElementById('btnClosePrivacyModal');

  const completeModal = document.getElementById('completeModal');
  const btnCloseCompleteModal = document.getElementById('btnCloseCompleteModal');
  const btnNewOrder = document.getElementById('btnNewOrder');
  const btnCopyReceipt = document.getElementById('btnCopyReceipt');
  const btnDownloadExcelModal = document.getElementById('btnDownloadExcelModal');
  const btnDownloadCsvModal = document.getElementById('btnDownloadCsvModal');

  // 2. 가상 도로명 주소 데이터베이스
  const ADDRESS_DATABASE = [
    { zip: '06164', road: '서울특별시 강남구 테헤란로 152', jibun: '역삼동 737 (강남파이낸스센터)' },
    { zip: '06236', road: '서울특별시 강남구 테헤란로 142', jibun: '역삼동 736-1 (아크플레이스)' },
    { zip: '04781', road: '서울특별시 성동구 성수이로 88', jibun: '성수동2가 315-1 (성수디벨로퍼스센터)' },
    { zip: '04790', road: '서울특별시 성동구 뚝섬로 273', jibun: '성수동1가 685-700 (갤러리아포레)' },
    { zip: '13494', road: '경기도 성남시 분당구 판교역로 166', jibun: '백현동 532 (카카오 판교아지트)' },
    { zip: '13487', road: '경기도 성남시 분당구 판교역로 235', jibun: '삼평동 681 (에이치스퀘어)' },
    { zip: '05551', road: '서울특별시 송파구 올림픽로 300', jibun: '신천동 29 (롯데월드타워)' },
    { zip: '04038', road: '서울특별시 마포구 양화로 160', jibun: '동교동 165-8 (홍대입구역 복합역사)' },
    { zip: '48058', road: '부산광역시 해운대구 센텀중앙로 78', jibun: '우동 1466-1 (센텀그린타워)' }
  ];

  // 3. 상품 장바구니 상태 및 가격 계산
  const cartItems = [
    { id: 1, name: '프렌치 워시드 리넨 셔츠', price: 79000, qty: 1 },
    { id: 2, name: '아틀리에 매트 세라믹 머그', price: 24000, qty: 2 }
  ];
  const FIRST_ORDER_DISCOUNT = 5000;

  function updateOrderTotals() {
    let subtotal = 0;
    cartItems.forEach((item) => {
      subtotal += item.price * item.qty;
    });

    const shippingFee = subtotal >= 50000 ? 0 : 3000;
    const finalTotal = Math.max(0, subtotal + shippingFee - FIRST_ORDER_DISCOUNT);

    document.getElementById('subtotalVal').textContent = subtotal.toLocaleString() + '원';
    document.getElementById('shippingVal').textContent =
      shippingFee === 0 ? '0원 (무료배송)' : shippingFee.toLocaleString() + '원';
    document.getElementById('finalTotalVal').textContent = finalTotal.toLocaleString() + '원';

    if (btnSubmit) {
      btnSubmit.textContent = `${finalTotal.toLocaleString()}원 결제 및 주문 접수하기`;
    }
  }

  // 장바구니 수량 조절
  document.querySelectorAll('.qty-btn').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      const target = e.currentTarget;
      const itemId = parseInt(target.getAttribute('data-id'), 10);
      const action = target.getAttribute('data-action');
      const item = cartItems.find((i) => i.id === itemId);

      if (!item) return;

      if (action === 'plus') {
        item.qty += 1;
      } else if (action === 'minus' && item.qty > 1) {
        item.qty -= 1;
      }

      const qtySpan = document.getElementById(`qty-${itemId}`);
      if (qtySpan) qtySpan.textContent = item.qty;

      const priceSpan = document.getElementById(`item-price-${itemId}`);
      if (priceSpan) priceSpan.textContent = (item.price * item.qty).toLocaleString() + '원';

      updateOrderTotals();
    });
  });

  // 4. 전화번호 자동 하이픈 포맷팅
  function formatPhoneNumber(inputElement) {
    let val = inputElement.value.replace(/[^0-9]/g, '');
    if (val.length > 11) val = val.slice(0, 11);

    if (val.length > 3 && val.length <= 7) {
      val = val.slice(0, 3) + '-' + val.slice(3);
    } else if (val.length > 7) {
      val = val.slice(0, 3) + '-' + val.slice(3, 7) + '-' + val.slice(7);
    }
    inputElement.value = val;
  }

  phoneInput.addEventListener('input', () => {
    formatPhoneNumber(phoneInput);
    validatePhone();
  });

  recipientPhoneInput.addEventListener('input', () => {
    formatPhoneNumber(recipientPhoneInput);
    validateRecipientPhone();
  });

  // 5. 이메일 도메인 선택
  domainSelect.addEventListener('change', () => {
    const selected = domainSelect.value;
    let currentEmail = emailInput.value.trim();

    if (selected === 'direct') {
      emailInput.placeholder = 'user@example.com';
      emailInput.focus();
    } else {
      const parts = currentEmail.split('@');
      const prefix = parts[0] || '';
      emailInput.value = prefix ? `${prefix}@${selected}` : `@${selected}`;
    }
    validateEmail();
  });

  // 6. 주문자와 수령인 동일 여부 토글
  sameAsRecipientCheck.addEventListener('change', () => {
    if (sameAsRecipientCheck.checked) {
      recipientGroup.classList.remove('active');
    } else {
      recipientGroup.classList.add('active');
      recipientNameInput.focus();
    }
  });

  // 7. 배송 요청사항 직접입력 토글
  deliverySelect.addEventListener('change', () => {
    if (deliverySelect.value === 'custom') {
      customDeliveryInput.style.display = 'block';
      customDeliveryInput.focus();
    } else {
      customDeliveryInput.style.display = 'none';
      customDeliveryInput.value = '';
    }
  });

  // 8. 유효성 검사 로직
  function showError(input, errorEl, message) {
    if (input) input.classList.add('is-invalid');
    if (errorEl) {
      errorEl.textContent = message;
      errorEl.classList.add('show');
    }
  }

  function clearError(input, errorEl) {
    if (input) input.classList.remove('is-invalid');
    if (errorEl) {
      errorEl.textContent = '';
      errorEl.classList.remove('show');
    }
  }

  function validateName() {
    const val = nameInput.value.trim();
    if (!val) {
      showError(nameInput, nameError, '이름을 입력해 주세요.');
      return false;
    }
    if (val.length < 2) {
      showError(nameInput, nameError, '이름은 최소 2글자 이상 입력해 주세요.');
      return false;
    }
    clearError(nameInput, nameError);
    return true;
  }

  function validatePhone() {
    const val = phoneInput.value.trim();
    const regex = /^01[016789]-?[0-9]{3,4}-?[0-9]{4}$/;
    if (!val) {
      showError(phoneInput, phoneError, '전화번호를 입력해 주세요.');
      return false;
    }
    if (!regex.test(val)) {
      showError(phoneInput, phoneError, '올바른 휴대폰 번호(010-XXXX-XXXX)를 입력해 주세요.');
      return false;
    }
    clearError(phoneInput, phoneError);
    return true;
  }

  function validateEmail() {
    const val = emailInput.value.trim();
    const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!val) {
      showError(emailInput, emailError, '이메일 주소를 입력해 주세요.');
      return false;
    }
    if (!regex.test(val)) {
      showError(emailInput, emailError, '올바른 이메일 형식(예: user@example.com)을 입력해 주세요.');
      return false;
    }
    clearError(emailInput, emailError);
    return true;
  }

  function validateAddress() {
    const zip = zipcodeInput.value.trim();
    const road = roadAddressInput.value.trim();
    const detail = detailAddressInput.value.trim();

    if (!zip || !road) {
      showError(detailAddressInput, addressError, '우편번호 찾기를 통해 기본 주소를 먼저 검색해 주세요.');
      return false;
    }
    if (!detail) {
      showError(detailAddressInput, addressError, '동/호수 등 상세 주소를 입력해 주세요.');
      return false;
    }
    clearError(detailAddressInput, addressError);
    return true;
  }

  function validateRecipientName() {
    if (sameAsRecipientCheck.checked) return true;
    const val = recipientNameInput.value.trim();
    if (!val || val.length < 2) {
      showError(recipientNameInput, recipientNameError, '수령인 이름을 2자 이상 입력해 주세요.');
      return false;
    }
    clearError(recipientNameInput, recipientNameError);
    return true;
  }

  function validateRecipientPhone() {
    if (sameAsRecipientCheck.checked) return true;
    const val = recipientPhoneInput.value.trim();
    const regex = /^01[016789]-?[0-9]{3,4}-?[0-9]{4}$/;
    if (!val || !regex.test(val)) {
      showError(recipientPhoneInput, recipientPhoneError, '수령인의 올바른 휴대폰 번호를 입력해 주세요.');
      return false;
    }
    clearError(recipientPhoneInput, recipientPhoneError);
    return true;
  }

  function validateTerms() {
    let isValid = true;
    if (!agreePrivacyCheck.checked) {
      showError(null, privacyError, '개인정보 수집 및 이용 동의가 필요합니다.');
      isValid = false;
    } else {
      clearError(null, privacyError);
    }

    if (!agreeOrderCheck.checked) {
      showError(null, orderError, '주문 상품 및 배송 정보 확인 동의가 필요합니다.');
      isValid = false;
    } else {
      clearError(null, orderError);
    }
    return isValid;
  }

  // 실시간 포커스 아웃 검사
  nameInput.addEventListener('blur', validateName);
  emailInput.addEventListener('blur', validateEmail);
  detailAddressInput.addEventListener('blur', validateAddress);
  agreePrivacyCheck.addEventListener('change', validateTerms);
  agreeOrderCheck.addEventListener('change', validateTerms);

  // 9. 우편번호 검색 모달
  function renderAddressResults(list) {
    addressResultList.innerHTML = '';
    if (list.length === 0) {
      addressResultList.innerHTML = `
        <div style="padding: 24px; text-align: center; color: var(--color-text-muted); font-size: 13px;">
          검색 결과가 없습니다. 도로명 또는 건물명을 확인해 주세요.
        </div>
      `;
      return;
    }

    list.forEach((item) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'address-item-btn';
      btn.innerHTML = `
        <div>
          <div class="addr-road">${item.road}</div>
          <div class="addr-jibun">${item.jibun}</div>
        </div>
        <div class="addr-zip">${item.zip}</div>
      `;
      btn.addEventListener('click', () => {
        zipcodeInput.value = item.zip;
        roadAddressInput.value = item.road;
        addressModal.classList.remove('open');
        clearError(detailAddressInput, addressError);
        detailAddressInput.focus();
      });
      addressResultList.appendChild(btn);
    });
  }

  btnSearchAddress.addEventListener('click', () => {
    addressModal.classList.add('open');
    addressSearchInput.value = '';
    renderAddressResults(ADDRESS_DATABASE);
    addressSearchInput.focus();
  });

  roadAddressInput.addEventListener('click', () => {
    btnSearchAddress.click();
  });

  addressSearchInput.addEventListener('input', (e) => {
    const query = e.target.value.trim().toLowerCase();
    if (!query) {
      renderAddressResults(ADDRESS_DATABASE);
      return;
    }
    const filtered = ADDRESS_DATABASE.filter(
      (item) =>
        item.road.toLowerCase().includes(query) ||
        item.jibun.toLowerCase().includes(query) ||
        item.zip.includes(query)
    );
    renderAddressResults(filtered);
  });

  document.querySelectorAll('.btn-tag-search').forEach((tag) => {
    tag.addEventListener('click', (e) => {
      const keyword = e.target.textContent.trim();
      addressSearchInput.value = keyword;
      addressSearchInput.dispatchEvent(new Event('input'));
    });
  });

  btnCloseAddressModal.addEventListener('click', () => {
    addressModal.classList.remove('open');
  });

  // 10. 개인정보 약관 모달
  btnOpenPrivacy.addEventListener('click', () => {
    privacyModal.classList.add('open');
  });
  btnClosePrivacyModal.addEventListener('click', () => {
    privacyModal.classList.remove('open');
  });

  // 11. 예시 데이터 자동 입력 (Demo Fill)
  btnFillDemo.addEventListener('click', () => {
    nameInput.value = '홍길동';
    phoneInput.value = '010-9876-5432';
    emailInput.value = 'gildong.hong@example.com';
    domainSelect.value = 'direct';

    zipcodeInput.value = '06164';
    roadAddressInput.value = '서울특별시 강남구 테헤란로 152';
    detailAddressInput.value = '강남파이낸스센터 18층 아틀리에오피스';

    sameAsRecipientCheck.checked = true;
    recipientGroup.classList.remove('active');

    deliverySelect.value = '부재 시 문 앞에 놓아주세요';
    customDeliveryInput.style.display = 'none';

    agreePrivacyCheck.checked = true;
    agreeOrderCheck.checked = true;

    document.querySelectorAll('.error-hint').forEach((el) => el.classList.remove('show'));
    document.querySelectorAll('.is-invalid').forEach((el) => el.classList.remove('is-invalid'));
  });

  // 12. 전체 초기화 (Reset)
  btnReset.addEventListener('click', () => {
    if (confirm('작성 중인 사용자 정보를 모두 초기화하시겠습니까?')) {
      form.reset();
      recipientGroup.classList.remove('active');
      customDeliveryInput.style.display = 'none';
      document.querySelectorAll('.error-hint').forEach((el) => el.classList.remove('show'));
      document.querySelectorAll('.is-invalid').forEach((el) => el.classList.remove('is-invalid'));
      updateOrderTotals();
    }
  });

  // 13. 입력 정보 수집 함수 (Helper)
  function getOrderFormData() {
    const orderNum = 'ORD-' + new Date().toISOString().slice(0, 10).replace(/-/g, '') + '-' + Math.floor(1000 + Math.random() * 9000);
    const orderDate = new Date().toLocaleString('ko-KR', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });
    const recipientName = sameAsRecipientCheck.checked ? nameInput.value.trim() : recipientNameInput.value.trim();
    const recipientPhone = sameAsRecipientCheck.checked ? phoneInput.value.trim() : recipientPhoneInput.value.trim();
    const memo = deliverySelect.value === 'custom' ? customDeliveryInput.value : deliverySelect.value;
    const itemsSummary = cartItems.map((i) => `${i.name}(${i.qty}개)`).join(', ');
    const subtotalText = document.getElementById('subtotalVal').textContent;
    const shippingText = document.getElementById('shippingVal').textContent;
    const totalText = document.getElementById('finalTotalVal').textContent;

    return {
      orderNum,
      orderDate,
      userName: nameInput.value.trim(),
      userPhone: phoneInput.value.trim(),
      userEmail: emailInput.value.trim(),
      recipient: `${recipientName} (${recipientPhone})`,
      recipientName,
      recipientPhone,
      zipcode: zipcodeInput.value.trim(),
      roadAddress: roadAddressInput.value.trim(),
      detailAddress: detailAddressInput.value.trim(),
      fullAddress: `(${zipcodeInput.value.trim()}) ${roadAddressInput.value.trim()} ${detailAddressInput.value.trim()}`,
      deliveryMemo: memo || '요청사항 없음',
      itemsSummary,
      subtotal: subtotalText,
      shippingFee: shippingText,
      finalTotal: totalText
    };
  }

  // 14. 엑셀 파일 저장 핵심 로직 (Excel XML Spreadsheet .xls & CSV with UTF-8 BOM)
  function triggerFileDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // A. 엑셀 파일 (.xls - Excel에서 서식 및 표가 완벽히 렌더링되는 공식 XML/HTML 포맷)
  function exportToExcelFile(data) {
    const filename = `주문자정보_${data.userName || '고객'}_${new Date().toISOString().slice(0, 10)}.xls`;
    const excelContent = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
        <!--[if gte mso 9]>
        <xml>
          <x:ExcelWorkbook>
            <x:ExcelWorksheets>
              <x:ExcelWorksheet>
                <x:Name>주문자_배송정보</x:Name>
                <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
              </x:ExcelWorksheet>
            </x:ExcelWorksheets>
          </x:ExcelWorkbook>
        </xml>
        <![endif]-->
        <style>
          table { border-collapse: collapse; width: 100%; font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; }
          th { background-color: #1e293b; color: #ffffff; border: 1px solid #94a3b8; padding: 10px; font-weight: bold; text-align: center; }
          td { border: 1px solid #cbd5e1; padding: 8px 12px; font-size: 13px; vertical-align: middle; }
          .header-title { font-size: 16px; font-weight: bold; height: 40px; background-color: #f1f5f9; text-align: center; }
          .sec-label { font-weight: bold; background-color: #f8fafc; text-align: center; }
          .text-center { text-align: center; }
          .text-right { text-align: right; }
          .mso-text { mso-number-format: "\\@"; }
          .total-highlight { font-weight: bold; color: #0f172a; font-size: 14px; background-color: #fef08a; }
        </style>
      </head>
      <body>
        <table>
          <tr>
            <th colspan="4" class="header-title">온라인 쇼핑몰 주문자 및 배송 정보 명세서</th>
          </tr>
          <tr>
            <td colspan="4" style="text-align: right; color: #64748b; font-size: 11px;">
              다운로드 일시: ${data.orderDate}
            </td>
          </tr>
          <tr>
            <th style="width: 120px;">대분류</th>
            <th style="width: 140px;">항목명</th>
            <th style="width: 380px;">고객 입력 내용</th>
            <th style="width: 100px;">비고</th>
          </tr>
          <tr>
            <td rowspan="2" class="sec-label">주문 정보</td>
            <td>주문번호</td>
            <td class="mso-text">${data.orderNum}</td>
            <td class="text-center">자동생성</td>
          </tr>
          <tr>
            <td>주문일시</td>
            <td>${data.orderDate}</td>
            <td class="text-center">접수완료</td>
          </tr>
          <tr>
            <td rowspan="3" class="sec-label">주문자 정보</td>
            <td>주문자 이름</td>
            <td style="font-weight: bold;">${data.userName}</td>
            <td class="text-center">필수</td>
          </tr>
          <tr>
            <td>전화번호</td>
            <td class="mso-text">${data.userPhone}</td>
            <td class="text-center">필수</td>
          </tr>
          <tr>
            <td>이메일</td>
            <td class="mso-text">${data.userEmail}</td>
            <td class="text-center">필수</td>
          </tr>
          <tr>
            <td rowspan="5" class="sec-label">배송지 정보</td>
            <td>수령인 정보</td>
            <td>${data.recipient}</td>
            <td class="text-center">수령인</td>
          </tr>
          <tr>
            <td>우편번호</td>
            <td class="mso-text">${data.zipcode}</td>
            <td class="text-center">5자리</td>
          </tr>
          <tr>
            <td>기본 주소</td>
            <td>${data.roadAddress}</td>
            <td class="text-center">도로명</td>
          </tr>
          <tr>
            <td>상세 주소</td>
            <td>${data.detailAddress}</td>
            <td class="text-center">동/호수</td>
          </tr>
          <tr>
            <td>배송 요청사항</td>
            <td>${data.deliveryMemo}</td>
            <td class="text-center">배송메모</td>
          </tr>
          <tr>
            <td rowspan="3" class="sec-label">결제 요약</td>
            <td>주문 상품</td>
            <td>${data.itemsSummary}</td>
            <td class="text-center">상품</td>
          </tr>
          <tr>
            <td>배송비</td>
            <td>${data.shippingFee}</td>
            <td class="text-center">무료</td>
          </tr>
          <tr>
            <td>최종 결제 금액</td>
            <td class="total-highlight text-right">${data.finalTotal}</td>
            <td class="text-center">결제예정</td>
          </tr>
        </table>
      </body>
      </html>
    `;

    const blob = new Blob([excelContent], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    triggerFileDownload(blob, filename);
  }

  // B. 엑셀 호환 CSV 파일 (UTF-8 BOM 포함)
  function exportToCsvFile(data) {
    const filename = `주문자정보_${data.userName || '고객'}_${new Date().toISOString().slice(0, 10)}.csv`;
    const BOM = '\uFEFF';
    const rows = [
      ['구분', '항목명', '고객 입력 내용', '비고'],
      ['주문정보', '주문번호', data.orderNum, '자동생성'],
      ['주문정보', '주문일시', data.orderDate, '접수일시'],
      ['주문자', '이름', data.userName, '필수'],
      ['주문자', '전화번호', data.userPhone, '필수'],
      ['주문자', '이메일', data.userEmail, '필수'],
      ['배송지', '수령인', data.recipient, ''],
      ['배송지', '우편번호', data.zipcode, '필수'],
      ['배송지', '기본주소', data.roadAddress, '필수'],
      ['배송지', '상세주소', data.detailAddress, '필수'],
      ['배송지', '배송요청사항', data.deliveryMemo, ''],
      ['결제정보', '주문상품목록', data.itemsSummary, ''],
      ['결제정보', '배송비', data.shippingFee, ''],
      ['결제정보', '최종결제금액', data.finalTotal, '']
    ];

    const csvBody = rows
      .map((row) => row.map((cell) => `"${String(cell || '').replace(/"/g, '""')}"`).join(','))
      .join('\r\n');

    const blob = new Blob([BOM + csvBody], { type: 'text/csv;charset=utf-8;' });
    triggerFileDownload(blob, filename);
  }

  // 15. 메인 폼의 [엑셀 파일로 저장] 버튼 핸들러
  btnExportExcelDirect.addEventListener('click', () => {
    const isNameOk = validateName();
    const isPhoneOk = validatePhone();
    const isEmailOk = validateEmail();
    const isAddrOk = validateAddress();

    if (!isNameOk || !isPhoneOk || !isEmailOk || !isAddrOk) {
      alert('엑셀 파일 저장을 위해 이름, 전화번호, 이메일, 주소를 먼저 정확히 입력해 주세요.');
      const firstInvalid = document.querySelector('.is-invalid, .error-hint.show');
      if (firstInvalid) {
        firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
      return;
    }

    const data = getOrderFormData();
    exportToExcelFile(data);
  });

  // 16. 메인 폼 제출 시 주문 완료 영수증 모달 출력
  form.addEventListener('submit', (e) => {
    e.preventDefault();

    const isNameOk = validateName();
    const isPhoneOk = validatePhone();
    const isEmailOk = validateEmail();
    const isAddrOk = validateAddress();
    const isRecipNameOk = validateRecipientName();
    const isRecipPhoneOk = validateRecipientPhone();
    const isTermsOk = validateTerms();

    if (isNameOk && isPhoneOk && isEmailOk && isAddrOk && isRecipNameOk && isRecipPhoneOk && isTermsOk) {
      const data = getOrderFormData();

      document.getElementById('receiptOrderNum').textContent = data.orderNum;
      document.getElementById('receiptName').textContent = data.userName;
      document.getElementById('receiptPhone').textContent = data.userPhone;
      document.getElementById('receiptEmail').textContent = data.userEmail;
      document.getElementById('receiptRecipient').textContent = data.recipient;
      document.getElementById('receiptAddress').textContent = data.fullAddress;
      document.getElementById('receiptMemo').textContent = data.deliveryMemo;
      document.getElementById('receiptAmount').textContent = data.finalTotal;

      completeModal.classList.add('open');
    } else {
      const firstInvalid = document.querySelector('.is-invalid, .error-hint.show');
      if (firstInvalid) {
        firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }
  });

  // 모달 안의 엑셀 다운로드 버튼들
  btnDownloadExcelModal.addEventListener('click', () => {
    const data = getOrderFormData();
    exportToExcelFile(data);
  });

  btnDownloadCsvModal.addEventListener('click', () => {
    const data = getOrderFormData();
    exportToCsvFile(data);
  });

  // 영수증 복사
  btnCopyReceipt.addEventListener('click', () => {
    const data = getOrderFormData();
    const text = `[ATELIER SEOUL 주문서]
주문번호: ${data.orderNum}
주문일시: ${data.orderDate}
주문자: ${data.userName} (${data.userPhone})
이메일: ${data.userEmail}
수령인: ${data.recipient}
배송지: ${data.fullAddress}
배송요청사항: ${data.deliveryMemo}
최종결제금액: ${data.finalTotal}`;

    navigator.clipboard.writeText(text).then(() => {
      alert('주문 내역 영수증이 클립보드에 복사되었습니다.');
    });
  });

  btnCloseCompleteModal.addEventListener('click', () => {
    completeModal.classList.remove('open');
  });

  btnNewOrder.addEventListener('click', () => {
    completeModal.classList.remove('open');
    form.reset();
    recipientGroup.classList.remove('active');
    customDeliveryInput.style.display = 'none';
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // 모달 닫기 이벤트들 (배경 클릭, ESC 키)
  [addressModal, privacyModal, completeModal].forEach((modal) => {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.classList.remove('open');
      }
    });
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      [addressModal, privacyModal, completeModal].forEach((modal) => {
        modal.classList.remove('open');
      });
    }
  });

  // 초기 주문 금액 계산 실행
  updateOrderTotals();
});

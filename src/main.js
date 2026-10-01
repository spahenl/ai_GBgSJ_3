/**
 * ATELIER SEOUL - 온라인 쇼핑몰 사용자 정보 입력 폼
 * Pure Vanilla JavaScript (순수 자바스크립트 구현)
 * + 엑셀(.xlsx) 파일 생성 및 저장 기능 (SheetJS)
 */
import * as XLSX from 'xlsx';

document.addEventListener('DOMContentLoaded', () => {
  // 1. 요소 참조 (DOM Elements)
  const form = document.getElementById('orderForm');
  const btnSubmit = document.getElementById('btnSubmit');
  const btnFillDemo = document.getElementById('btnFillDemo');
  const btnReset = document.getElementById('btnReset');
  const btnExportExcelHeader = document.getElementById('btnExportExcelHeader');
  const btnQuickExcel = document.getElementById('btnQuickExcel');
  const btnDownloadExcel = document.getElementById('btnDownloadExcel');

  // 사용자 정보 입력 필드
  const nameInput = document.getElementById('userName');
  const phoneInput = document.getElementById('userPhone');
  const emailInput = document.getElementById('userEmail');
  const domainSelect = document.getElementById('emailDomainSelect');
  
  // 배송지 입력 필드
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

  // 토스트 알림
  const excelToast = document.getElementById('excelToast');
  const excelToastMsg = document.getElementById('excelToastMsg');

  let latestSubmittedOrder = null;

  // 2. 가상 주소 데이터베이스 (주소 검색용)
  const ADDRESS_DATABASE = [
    {
      zip: '06164',
      road: '서울특별시 강남구 테헤란로 152',
      jibun: '서울특별시 강남구 역삼동 737 (강남파이낸스센터)'
    },
    {
      zip: '06236',
      road: '서울특별시 강남구 테헤란로 142',
      jibun: '서울특별시 강남구 역삼동 736-1 (아크플레이스)'
    },
    {
      zip: '04781',
      road: '서울특별시 성동구 성수이로 88',
      jibun: '서울특별시 성동구 성수동2가 315-1 (성수디벨로퍼스센터)'
    },
    {
      zip: '04790',
      road: '서울특별시 성동구 뚝섬로 273',
      jibun: '서울특별시 성동구 성수동1가 685-700 (갤러리아포레)'
    },
    {
      zip: '13494',
      road: '경기도 성남시 분당구 판교역로 166',
      jibun: '경기도 성남시 분당구 백현동 532 (카카오 판교아지트)'
    },
    {
      zip: '13487',
      road: '경기도 성남시 분당구 판교역로 235',
      jibun: '경기도 성남시 분당구 삼평동 681 (에이치스퀘어)'
    },
    {
      zip: '05551',
      road: '서울특별시 송파구 올림픽로 300',
      jibun: '서울특별시 송파구 신천동 29 (롯데월드타워)'
    },
    {
      zip: '04038',
      road: '서울특별시 마포구 양화로 160',
      jibun: '서울특별시 마포구 동교동 165-8 (홍대입구역 복합역사)'
    },
    {
      zip: '48058',
      road: '부산광역시 해운대구 센텀중앙로 78',
      jibun: '부산광역시 해운대구 우동 1466-1 (센텀그린타워)'
    }
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

    // DOM 업데이트
    document.getElementById('subtotalVal').textContent = subtotal.toLocaleString() + '원';
    document.getElementById('shippingVal').textContent =
      shippingFee === 0 ? '0원 (무료배송)' : shippingFee.toLocaleString() + '원';
    document.getElementById('finalTotalVal').textContent = finalTotal.toLocaleString() + '원';

    if (btnSubmit) {
      btnSubmit.textContent = finalTotal.toLocaleString() + '원 결제 및 주문 접수하기';
    }
  }

  // 장바구니 수량 조절 버튼 바인딩
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

  // 4. 전화번호 자동 하이픈 (-) 포맷터
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

  // 5. 이메일 도메인 간편 선택 연동
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

  // 8. 유효성 검사 함수들 (Validation Rules)
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

  // 실시간 포커스 아웃 이벤트 연결
  nameInput.addEventListener('blur', validateName);
  emailInput.addEventListener('blur', validateEmail);
  detailAddressInput.addEventListener('blur', validateAddress);
  agreePrivacyCheck.addEventListener('change', validateTerms);
  agreeOrderCheck.addEventListener('change', validateTerms);

  // 9. 우편번호 검색 모달 로직
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

  // 추천 검색어 클릭 지원
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
    emailInput.value = 'gildong.hong@atelier.com';
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

    // 기존 에러들 초기화
    document.querySelectorAll('.error-hint').forEach((el) => el.classList.remove('show'));
    document.querySelectorAll('.is-invalid').forEach((el) => el.classList.remove('is-invalid'));

    showToast('예시 사용자 정보가 채워졌습니다. 이제 엑셀로 저장하거나 결제를 진행해 보세요.');
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
      showToast('입력 내용이 초기화되었습니다.');
    }
  });

  // =========================================================================
  // 13. 엑셀(.xlsx) 파일 생성 및 다운로드 로직 (SheetJS)
  // =========================================================================
  function showToast(message) {
    if (!excelToast || !excelToastMsg) return;
    excelToastMsg.textContent = message;
    excelToast.classList.add('show');
    setTimeout(() => {
      excelToast.classList.remove('show');
    }, 3200);
  }

  function getFormDataSnapshot() {
    const recipientName = sameAsRecipientCheck.checked ? nameInput.value.trim() : recipientNameInput.value.trim();
    const recipientPhone = sameAsRecipientCheck.checked ? phoneInput.value.trim() : recipientPhoneInput.value.trim();
    const fullAddress = zipcodeInput.value.trim() || roadAddressInput.value.trim()
      ? `(${zipcodeInput.value.trim()}) ${roadAddressInput.value.trim()} ${detailAddressInput.value.trim()}`.trim()
      : '';
    const memo = deliverySelect.value === 'custom' ? customDeliveryInput.value.trim() : deliverySelect.value;
    const itemsSummary = cartItems.map((i) => `${i.name}(${i.qty}개)`).join(', ');
    const totalAmount = document.getElementById('finalTotalVal').textContent;

    return {
      orderNum: 'ORD-' + new Date().toISOString().slice(0, 10).replace(/-/g, '') + '-' + Math.floor(1000 + Math.random() * 9000),
      orderDate: new Date().toLocaleString('ko-KR'),
      name: nameInput.value.trim(),
      phone: phoneInput.value.trim(),
      email: emailInput.value.trim(),
      recipientName: recipientName || nameInput.value.trim(),
      recipientPhone: recipientPhone || phoneInput.value.trim(),
      zipcode: zipcodeInput.value.trim(),
      roadAddress: roadAddressInput.value.trim(),
      detailAddress: detailAddressInput.value.trim(),
      fullAddress: fullAddress,
      deliveryMemo: memo || '부재 시 문 앞에 놓아주세요',
      itemsSummary: itemsSummary,
      totalAmount: totalAmount,
      agreePrivacy: agreePrivacyCheck.checked ? '동의' : '미동의',
      agreeOrder: agreeOrderCheck.checked ? '동의' : '미동의'
    };
  }

  function saveToExcel(orderData, filenamePrefix) {
    const headers = [
      '주문번호',
      '주문일시',
      '주문자명',
      '전화번호',
      '이메일',
      '수령인명',
      '수령인연락처',
      '우편번호',
      '기본주소',
      '상세주소',
      '전체주소',
      '배송요청사항',
      '주문상품목록',
      '최종결제금액',
      '개인정보수집동의'
    ];

    const row = [
      orderData.orderNum || '',
      orderData.orderDate || new Date().toLocaleString('ko-KR'),
      orderData.name || '',
      orderData.phone || '',
      orderData.email || '',
      orderData.recipientName || '',
      orderData.recipientPhone || '',
      orderData.zipcode || '',
      orderData.roadAddress || '',
      orderData.detailAddress || '',
      orderData.fullAddress || '',
      orderData.deliveryMemo || '',
      orderData.itemsSummary || '',
      orderData.totalAmount || '',
      orderData.agreePrivacy || ''
    ];

    const worksheetData = [headers, row];
    const ws = XLSX.utils.aoa_to_sheet(worksheetData);

    // 엑셀 컬럼 너비(Width) 자동 설정
    ws['!cols'] = [
      { wch: 20 }, // 주문번호
      { wch: 22 }, // 주문일시
      { wch: 14 }, // 주문자명
      { wch: 16 }, // 전화번호
      { wch: 25 }, // 이메일
      { wch: 14 }, // 수령인명
      { wch: 16 }, // 수령인연락처
      { wch: 10 }, // 우편번호
      { wch: 32 }, // 기본주소
      { wch: 25 }, // 상세주소
      { wch: 45 }, // 전체주소
      { wch: 26 }, // 배송요청사항
      { wch: 35 }, // 주문상품목록
      { wch: 15 }, // 최종결제금액
      { wch: 14 }  // 개인정보수집동의
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '주문_고객정보');

    const customerName = orderData.name ? `_${orderData.name}` : '';
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const filename = `${filenamePrefix}${customerName}_${dateStr}.xlsx`;

    XLSX.writeFile(wb, filename);
    showToast(`엑셀 파일(${filename})이 다운로드되었습니다.`);
  }

  // 상단 헤더의 엑셀 다운로드 버튼
  if (btnExportExcelHeader) {
    btnExportExcelHeader.addEventListener('click', () => {
      const data = getFormDataSnapshot();
      if (!data.name && !data.phone && !data.email && !data.roadAddress) {
        if (confirm('아직 입력된 사용자 정보가 없습니다. 테스트용 예시 데이터를 자동으로 채운 뒤 엑셀로 저장할까요?')) {
          btnFillDemo.click();
          setTimeout(() => {
            const filledData = getFormDataSnapshot();
            saveToExcel(filledData, '온라인쇼핑몰_사용자정보');
          }, 100);
        }
        return;
      }
      saveToExcel(data, '온라인쇼핑몰_사용자정보');
    });
  }

  // 폼 하단의 작성 중 사용자 정보 즉시 저장 버튼
  if (btnQuickExcel) {
    btnQuickExcel.addEventListener('click', () => {
      const data = getFormDataSnapshot();
      if (!data.name && !data.phone && !data.email && !data.roadAddress) {
        alert('이름, 전화번호, 주소, 이메일 중 최소 하나 이상의 정보를 입력해 주세요.');
        nameInput.focus();
        return;
      }
      saveToExcel(data, '쇼핑몰_주문자정보');
    });
  }

  // 주문 완료 팝업 내부의 엑셀 다운로드 버튼
  if (btnDownloadExcel) {
    btnDownloadExcel.addEventListener('click', () => {
      if (latestSubmittedOrder) {
        saveToExcel(latestSubmittedOrder, '주문완료_영수증');
      } else {
        const data = getFormDataSnapshot();
        saveToExcel(data, '주문완료_영수증');
      }
    });
  }

  // 14. 폼 제출 (Submit) 및 주문 완료 영수증 출력
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
      const snapshot = getFormDataSnapshot();
      latestSubmittedOrder = snapshot;

      document.getElementById('receiptOrderNum').textContent = snapshot.orderNum;
      document.getElementById('receiptName').textContent = snapshot.name;
      document.getElementById('receiptPhone').textContent = snapshot.phone;
      document.getElementById('receiptEmail').textContent = snapshot.email;
      document.getElementById('receiptRecipient').textContent = `${snapshot.recipientName} (${snapshot.recipientPhone})`;
      document.getElementById('receiptAddress').textContent = snapshot.fullAddress;
      document.getElementById('receiptMemo').textContent = snapshot.deliveryMemo || '없음';
      document.getElementById('receiptAmount').textContent = snapshot.totalAmount;

      completeModal.classList.add('open');
      showToast('주문 정보가 접수되었습니다. 영수증 팝업에서 엑셀 파일 저장이 가능합니다.');
    } else {
      // 첫 번째 에러 요소로 스크롤
      const firstInvalid = document.querySelector('.is-invalid, .error-hint.show');
      if (firstInvalid) {
        firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }
  });

  // 주문 완료 닫기 및 새 주문
  btnCloseCompleteModal.addEventListener('click', () => {
    completeModal.classList.remove('open');
  });

  btnNewOrder.addEventListener('click', () => {
    completeModal.classList.remove('open');
    form.reset();
    recipientGroup.classList.remove('active');
    customDeliveryInput.style.display = 'none';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('새 주문 작성을 위해 양식이 초기화되었습니다.');
  });

  // 영수증 클립보드 복사
  btnCopyReceipt.addEventListener('click', () => {
    const text = `[ATELIER SEOUL 주문서]
주문번호: ${document.getElementById('receiptOrderNum').textContent}
주문자: ${document.getElementById('receiptName').textContent}
전화번호: ${document.getElementById('receiptPhone').textContent}
이메일: ${document.getElementById('receiptEmail').textContent}
받는 분: ${document.getElementById('receiptRecipient').textContent}
배송지: ${document.getElementById('receiptAddress').textContent}
배송메모: ${document.getElementById('receiptMemo').textContent}
결제금액: ${document.getElementById('receiptAmount').textContent}`;

    navigator.clipboard.writeText(text).then(() => {
      alert('주문 내역 영수증이 클립보드에 복사되었습니다.');
    });
  });

  // 바깥 배경 클릭 시 모달 닫기
  [addressModal, privacyModal, completeModal].forEach((modal) => {
    if (!modal) return;
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.classList.remove('open');
      }
    });
  });

  // ESC 키 누를 시 모달 닫기
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      [addressModal, privacyModal, completeModal].forEach((modal) => {
        if (modal) modal.classList.remove('open');
      });
    }
  });

  // 초기 주문 금액 계산 실행
  updateOrderTotals();
});

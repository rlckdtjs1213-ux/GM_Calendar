'use strict';

// 샘플 배포일정 엑셀 생성: npm run sample
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const outDir = path.join(__dirname, '..', 'sample-data');
fs.mkdirSync(outDir, { recursive: true });

const today = new Date();
function dayStr(offset) {
  const d = new Date(today);
  d.setDate(d.getDate() + offset);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

const rows = [
  { 날짜: dayStr(0), 제목: '주문 시스템 v2.3 배포', 담당자: '김철수', 비고: '운영 반영, 무중단' },
  { 날짜: dayStr(2), 제목: '결제 모듈 핫픽스', 담당자: '이영희', 비고: 'PG 연동 버그 수정' },
  { 날짜: dayStr(5), 제목: '회원 API 개편', 담당자: '박민수', 비고: '점검 22:00~23:00' },
  { 날짜: dayStr(7), 제목: '배치 서버 패치', 담당자: '김철수', 비고: '' },
  { 날짜: dayStr(12), 제목: '관리자 페이지 개편', 담당자: '최지은', 비고: '디자인 전면 변경' },
];

const ws = XLSX.utils.json_to_sheet(rows);
ws['!cols'] = [{ wch: 12 }, { wch: 28 }, { wch: 10 }, { wch: 30 }];
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '배포일정');

const outPath = path.join(outDir, '배포일정_샘플.xlsx');
XLSX.writeFile(wb, outPath);
console.log('샘플 생성 완료:', outPath);

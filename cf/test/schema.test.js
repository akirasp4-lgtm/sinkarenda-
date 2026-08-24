import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';

describe('schema.sql', () => {
  const sql = readFileSync(new URL('../schema.sql', import.meta.url), 'utf8');

  // --- 2026-08-24 最終総合レビュー（Fable 5 / Codex）で「切り替え不可」の判定を
  // 受け、D1の持ち方を「行ごとのテーブル」から「スナップショット1行」へ変更した。
  // 理由: (1)原子性 全件DELETE+分割INSERTでは途中状態が読めてしまう、
  //       (2)費用 1回の同期=5,765行×288回/日=166万行/日で無料枠(10万行/日)を超過。

  it('snapshotテーブルが必要な列（id/payload/hash/rows/bytes/at）を持つ', () => {
    const m = sql.match(/CREATE TABLE IF NOT EXISTS\s+snapshot\s*\([\s\S]*?\n\);/i);
    expect(m, 'snapshotのCREATE TABLE文が見つからない').toBeTruthy();
    for (const col of ['id', 'payload', 'hash', 'rows', 'bytes', 'at']) {
      expect(m[0]).toMatch(new RegExp('\\b' + col + '\\b'));
    }
  });

  it('snapshotのidは常に1固定（CHECK制約で複数行を防ぐ）', () => {
    const m = sql.match(/CREATE TABLE IF NOT EXISTS\s+snapshot\s*\([\s\S]*?\n\);/i);
    expect(m[0]).toMatch(/CHECK\s*\(\s*id\s*=\s*1\s*\)/i);
  });

  it('単価(rate)は保存しない（給料情報をD1へ持ち込まない）', () => {
    expect(sql).not.toMatch(/\brate\b/i);
  });

  it('sync_logテーブルを維持している（Cronの障害調査用）', () => {
    const m = sql.match(/CREATE TABLE IF NOT EXISTS\s+sync_log\s*\([\s\S]*?\n\);/i);
    expect(m, 'sync_logのCREATE TABLE文が見つからない').toBeTruthy();
    for (const col of ['at', 'rows', 'ok', 'message']) expect(m[0]).toMatch(new RegExp('\\b' + col + '\\b'));
  });

  it('sync_logはpayload_hash列を持つ（3回目レビュー修正3: 急減ガードの自己回復が同一内容かを判定するため）', () => {
    const m = sql.match(/CREATE TABLE IF NOT EXISTS\s+sync_log\s*\([\s\S]*?\n\);/i);
    expect(m[0]).toMatch(/\bpayload_hash\b/);
  });

  it('sync_logも列構成が変わったためDROPしてから作り直す（3回目レビュー修正3で列を追加したため）', () => {
    expect(sql).toMatch(/DROP TABLE IF EXISTS\s+sync_log/i);
  });

  it('sync_lockテーブルがある（修正2: 同時実行の抑止用）', () => {
    const m = sql.match(/CREATE TABLE IF NOT EXISTS\s+sync_lock\s*\([\s\S]*?\n\);/i);
    expect(m, 'sync_lockのCREATE TABLE文が見つからない').toBeTruthy();
    expect(m[0]).toMatch(/locked_at/i);
  });

  it('旧設計（行ごとのテーブル: nippo/members/genba/jobsites）は明示的にDROPしている', () => {
    for (const table of ['nippo', 'members', 'genba', 'jobsites']) {
      expect(sql).toMatch(new RegExp('DROP TABLE IF EXISTS\\s+' + table, 'i'));
    }
  });

  it('旧設計のCREATE TABLE文はもう残っていない', () => {
    for (const table of ['nippo', 'members', 'genba', 'jobsites']) {
      expect(sql).not.toMatch(new RegExp('CREATE TABLE IF NOT EXISTS\\s+' + table + '\\s*\\(', 'i'));
    }
  });
});

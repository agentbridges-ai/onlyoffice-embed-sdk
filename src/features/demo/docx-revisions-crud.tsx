"use client";

import { useState } from "react";
import type { RevisionItem } from "@/components/onlyoffice-embed-sdk";
import { DemoButton, DemoMenuRow } from "./demo-toolbar";
import type { MultiInstanceManager } from "./subframe-manager";

type DocxRevisionsCrudProps = {
  disabled?: boolean;
  getManager: () =>
    MultiInstanceManager | Promise<MultiInstanceManager | null> | null;
  onError?: (message: string, err: unknown) => void;
};

function revisionLabel(revision: RevisionItem | undefined) {
  if (!revision) return "";
  const type = revision.Data.TypeName || revision.Data.Type || "修订";
  const value = revision.Data.Value ? `：${revision.Data.Value}` : "";
  return `${type}${value}`;
}

export function DocxRevisionsCrud({
  disabled = false,
  getManager,
  onError,
}: DocxRevisionsCrudProps) {
  const [lastRevisionId, setLastRevisionId] = useState("");
  const [tracking, setTracking] = useState(false);
  const [status, setStatus] = useState("0 条");

  const runRevisionAction = async (
    action: (manager: MultiInstanceManager) => void | Promise<void>,
    errorMessage: string,
  ) => {
    try {
      const manager = await getManager();
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      await action(manager);
    } catch (err) {
      setStatus("失败");
      onError?.(errorMessage, err);
    }
  };

  const refreshRevisions = async (manager: MultiInstanceManager) => {
    const revisions = await manager.getEditor().getAllRevisions();
    setStatus(`${revisions.length} 条`);
    return revisions;
  };

  const toggleTrackRevisions = () =>
    runRevisionAction(async (manager) => {
      const next = !tracking;
      await manager.getEditor().setTrackRevisions(next);
      setTracking(next);
      setStatus(next ? "已开启追踪" : "已关闭追踪");
    }, "切换修订追踪失败");

  const addRevision = () =>
    runRevisionAction(async (manager) => {
      const revisions = await manager
        .getEditor()
        .addDemoRevision(`审批修订 ${new Date().toLocaleTimeString()}`);
      const latest = revisions.at(-1);
      if (latest) {
        setLastRevisionId(latest.Id);
      }
      setTracking(true);
      setStatus(`已新增 ${revisions.length} 条`);
    }, "新增审批修订失败");

  const readRevisions = () =>
    runRevisionAction(async (manager) => {
      const revisions = await refreshRevisions(manager);
      const latest = revisions.at(-1);
      if (latest) {
        setLastRevisionId(latest.Id);
        setStatus(`读取 ${revisions.length} 条：${revisionLabel(latest)}`);
      }
    }, "读取审批修订失败");

  const refreshAllRevisions = () =>
    runRevisionAction(async (manager) => {
      const revisions = await refreshRevisions(manager);
      console.log("[OnlyOffice demo] 全量审批修订", revisions);
      const latest = revisions.at(-1);
      if (latest) {
        setLastRevisionId(latest.Id);
      }
      setStatus(`全量 ${revisions.length} 条`);
    }, "全量读取审批修订失败");

  const goToRevision = () =>
    runRevisionAction(async (manager) => {
      const revisions = await refreshRevisions(manager);
      const id = lastRevisionId || revisions.at(-1)?.Id;
      if (!id) {
        setStatus("无可定位");
        return;
      }
      await manager.getEditor().goToRevision(id);
      setLastRevisionId(id);
      setStatus(`已定位 ${id}`);
    }, "定位审批修订失败");

  const acceptRevision = () =>
    runRevisionAction(async (manager) => {
      const revisions = await refreshRevisions(manager);
      const id = lastRevisionId || revisions.at(-1)?.Id;
      if (!id) {
        setStatus("无可接受");
        return;
      }
      await manager.getEditor().acceptRevision(id);
      setLastRevisionId("");
      const nextRevisions = await manager.getEditor().getAllRevisions();
      setStatus(`已接受，剩 ${nextRevisions.length} 条`);
    }, "接受审批修订失败");

  const rejectRevision = () =>
    runRevisionAction(async (manager) => {
      const revisions = await refreshRevisions(manager);
      const id = lastRevisionId || revisions.at(-1)?.Id;
      if (!id) {
        setStatus("无可拒绝");
        return;
      }
      await manager.getEditor().rejectRevision(id);
      setLastRevisionId("");
      const nextRevisions = await manager.getEditor().getAllRevisions();
      setStatus(`已拒绝，剩 ${nextRevisions.length} 条`);
    }, "拒绝审批修订失败");

  const acceptAllRevisions = () =>
    runRevisionAction(async (manager) => {
      await manager.getEditor().acceptAllRevisions();
      const revisions = await refreshRevisions(manager);
      setLastRevisionId("");
      setStatus(`已全部接受，剩 ${revisions.length} 条`);
    }, "接受全部审批修订失败");

  const rejectAllRevisions = () =>
    runRevisionAction(async (manager) => {
      await manager.getEditor().rejectAllRevisions();
      const revisions = await refreshRevisions(manager);
      setLastRevisionId("");
      setStatus(`已全部拒绝，剩 ${revisions.length} 条`);
    }, "拒绝全部审批修订失败");

  return (
    <DemoMenuRow>
      <span className="shrink-0 text-[13px] text-neutral-500">审批修订</span>
      <DemoButton
        disabled={disabled}
        active={tracking}
        onClick={toggleTrackRevisions}
      >
        {tracking ? "追踪中" : "追踪"}
      </DemoButton>
      <DemoButton disabled={disabled} onClick={addRevision}>
        新增
      </DemoButton>
      <DemoButton disabled={disabled} onClick={readRevisions}>
        读取
      </DemoButton>
      <DemoButton disabled={disabled} onClick={refreshAllRevisions}>
        全量
      </DemoButton>
      <DemoButton disabled={disabled} onClick={goToRevision}>
        定位
      </DemoButton>
      <DemoButton disabled={disabled} onClick={acceptRevision}>
        接受
      </DemoButton>
      <DemoButton disabled={disabled} onClick={rejectRevision}>
        拒绝
      </DemoButton>
      <DemoButton disabled={disabled} onClick={acceptAllRevisions}>
        全接
      </DemoButton>
      <DemoButton disabled={disabled} onClick={rejectAllRevisions}>
        全拒
      </DemoButton>
      <span
        className="max-w-[220px] truncate text-[12px] text-neutral-500"
        title={status}
      >
        {status}
      </span>
    </DemoMenuRow>
  );
}

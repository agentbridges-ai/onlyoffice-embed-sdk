"use client";

import { useState } from "react";
import { DemoButton, DemoMenuRow } from "./demo-toolbar";
import type { MultiInstanceManager } from "./subframe-manager";

type DocxCommentsCrudProps = {
  disabled?: boolean;
  getManager: () =>
    MultiInstanceManager | Promise<MultiInstanceManager | null> | null;
  onError?: (message: string, err: unknown) => void;
};

function getCommentText(data: unknown) {
  if (!data || typeof data !== "object") return "";

  const comment = data as { Text?: unknown; asc_getText?: () => unknown };
  const sdkText = comment.asc_getText?.();
  return String(comment.Text ?? sdkText ?? "");
}

export function DocxCommentsCrud({
  disabled = false,
  getManager,
  onError,
}: DocxCommentsCrudProps) {
  const [lastCommentId, setLastCommentId] = useState("");
  const [status, setStatus] = useState("0 条");

  const runCommentAction = async (
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

  const refreshCount = async (manager: MultiInstanceManager) => {
    const comments = await manager.getEditor().getAllComments();
    setStatus(`${comments.length} 条`);
    return comments;
  };

  const addComment = () =>
    runCommentAction(async (manager) => {
      const id = await manager.getEditor().addComment({
        Text: `审批批注 ${new Date().toLocaleTimeString()}`,
        UserName: "审批人",
        UserData: "approval-comment-crud",
      });
      setLastCommentId(id);
      const comments = await refreshCount(manager);
      setStatus(id ? `已新增 ${comments.length} 条` : "新增未返回 ID");
    }, "新增审批批注失败");

  const refreshAllComments = () =>
    runCommentAction(async (manager) => {
      const comments = await refreshCount(manager);
      console.log("[OnlyOffice demo] 全量审批批注", comments);
      const latest = comments.at(-1);
      if (latest) {
        setLastCommentId(latest.Id);
      }
      setStatus(`全量 ${comments.length} 条`);
    }, "全量读取审批批注失败");

  const readComments = () =>
    runCommentAction(async (manager) => {
      const comments = await refreshCount(manager);
      const latest = comments.at(-1);
      if (latest) {
        setLastCommentId(latest.Id);
        setStatus(`读取 ${comments.length} 条：${getCommentText(latest.Data)}`);
      }
    }, "读取审批批注失败");

  const updateComment = () =>
    runCommentAction(async (manager) => {
      const comments = await refreshCount(manager);
      const id = lastCommentId || comments.at(-1)?.Id;
      if (!id) {
        setStatus("无可修改");
        return;
      }

      await manager.getEditor().updateComment(id, {
        Text: `审批批注已修改 ${new Date().toLocaleTimeString()}`,
        UserName: "审批人",
        UserData: "approval-comment-crud",
      });
      setLastCommentId(id);
      setStatus(`已修改 ${id}`);
    }, "修改审批批注失败");

  const removeComment = () =>
    runCommentAction(async (manager) => {
      const comments = await refreshCount(manager);
      const id = lastCommentId || comments.at(-1)?.Id;
      if (!id) {
        setStatus("无可删除");
        return;
      }

      await manager.getEditor().removeComment(id);
      setLastCommentId("");
      const nextComments = await manager.getEditor().getAllComments();
      setStatus(`已删除，剩 ${nextComments.length} 条`);
    }, "删除审批批注失败");

  return (
    <DemoMenuRow>
      <span className="shrink-0 text-[13px] text-neutral-500">审批批注</span>
      <DemoButton disabled={disabled} onClick={addComment}>
        新增
      </DemoButton>
      <DemoButton disabled={disabled} onClick={readComments}>
        读取
      </DemoButton>
      <DemoButton disabled={disabled} onClick={refreshAllComments}>
        全量
      </DemoButton>
      <DemoButton disabled={disabled} onClick={updateComment}>
        修改
      </DemoButton>
      <DemoButton disabled={disabled} onClick={removeComment}>
        删除
      </DemoButton>
      <span
        className="max-w-[180px] truncate text-[12px] text-neutral-500"
        title={status}
      >
        {status}
      </span>
    </DemoMenuRow>
  );
}

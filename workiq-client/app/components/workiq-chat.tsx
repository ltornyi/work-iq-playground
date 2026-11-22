import {
  InteractionRequiredAuthError,
  type AccountInfo,
} from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { type FormEvent, useEffect, useRef, useState } from "react";

const WORK_IQ_SCOPE =
  "api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask";
const CONVERSATIONS_URL =
  "https://workiq.svc.cloud.microsoft/rest/beta/conversations";

export function WorkIqChatPanel({ account }: { account: AccountInfo }) {
  const { instance } = useMsal();
  const chatContainerRef = useRef<HTMLDivElement>(null);
  const shouldFollowLatestRef = useRef(true);
  const [conversationId, setConversationId] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isCreating, setIsCreating] = useState(false);
  const [isSending, setIsSending] = useState(false);
  const [prompt, setPrompt] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [pendingPrompt, setPendingPrompt] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const storageKey = `workiq.conversation.${account.homeAccountId}`;

  useEffect(() => {
    setIsLoading(true);
    setConversationId(null);
    setPrompt("");
    setMessages([]);
    setPendingPrompt(null);
    shouldFollowLatestRef.current = true;
    setErrorMessage(null);

    try {
      setConversationId(window.localStorage.getItem(storageKey));
    } catch (error) {
      setErrorMessage(
        `Couldn't read the saved Work IQ chat: ${getErrorMessage(error)}`,
      );
    } finally {
      setIsLoading(false);
    }
  }, [storageKey]);

  useEffect(() => {
    const container = chatContainerRef.current;
    if (container && shouldFollowLatestRef.current) {
      container.scrollTop = container.scrollHeight;
    }
  }, [messages, pendingPrompt, isSending]);

  const acquireAccessToken = async () => {
    const tokenRequest = {
      account,
      scopes: [WORK_IQ_SCOPE],
    };

    try {
      const { accessToken } = await instance.acquireTokenSilent(tokenRequest);
      return accessToken;
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) throw error;

      const { accessToken } = await instance.acquireTokenPopup(tokenRequest);
      return accessToken;
    }
  };

  const createChat = async () => {
    setIsCreating(true);
    setErrorMessage(null);

    try {
      const accessToken = await acquireAccessToken();
      const response = await fetch(CONVERSATIONS_URL, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: "{}",
      });

      if (!response.ok) {
        const details = await response.text();
        throw new Error(
          `Work IQ couldn't create a chat (${response.status})${
            details ? `: ${details}` : "."
          }`,
        );
      }

      const conversation = await readConversation(response);
      setConversationId(conversation.id);

      try {
        window.localStorage.setItem(storageKey, conversation.id);
      } catch (error) {
        setErrorMessage(
          `The chat was created, but couldn't be saved in this browser: ${getErrorMessage(error)}`,
        );
      }
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    } finally {
      setIsCreating(false);
    }
  };

  const sendPrompt = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    const message = prompt.trim();
    if (!conversationId || !message || isSending) return;

    shouldFollowLatestRef.current = true;
    setIsSending(true);
    setPendingPrompt(message);
    setPrompt("");
    setErrorMessage(null);

    try {
      const accessToken = await acquireAccessToken();
      const response = await fetch(
        `${CONVERSATIONS_URL}/${encodeURIComponent(conversationId)}/chat`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            message: {
              text: message,
            },
            locationHint: {
              timeZone:
                Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC",
            },
          }),
        },
      );

      if (!response.ok) {
        const details = await response.text();
        throw new Error(
          `Work IQ couldn't send the message (${response.status})${
            details ? `: ${details}` : "."
          }`,
        );
      }

      const conversation = await readChatConversation(response);
      setMessages(conversation.messages);
      setPendingPrompt(null);
    } catch (error) {
      setPrompt(message);
      setPendingPrompt(null);
      setErrorMessage(getErrorMessage(error));
    } finally {
      setIsSending(false);
    }
  };

  return (
    <section
      aria-labelledby="workiq-chat-heading"
      className="mt-6 rounded-2xl border border-white/10 bg-white/[0.03] p-4"
    >
      <h2
        className="text-sm font-semibold text-slate-100"
        id="workiq-chat-heading"
      >
        Work IQ chat
      </h2>

      {isLoading ? (
        <p aria-live="polite" className="mt-3 text-sm text-slate-400" role="status">
          Checking for a saved chat…
        </p>
      ) : conversationId ? (
        <div className="mt-3">
          <p className="inline-flex items-center gap-2 text-sm font-medium text-emerald-300">
            <span className="h-1.5 w-1.5 rounded-full bg-emerald-300" />
            Active chat
          </p>
          <p className="mt-2 text-xs text-slate-400">Conversation ID</p>
          <p className="mt-1 break-all font-mono text-sm text-slate-200">
            {conversationId}
          </p>
          <div
            aria-label="Conversation"
            aria-live="polite"
            aria-busy={isSending}
            className="mt-5 max-h-[min(32rem,60vh)] space-y-4 overflow-y-auto overscroll-contain rounded-2xl border border-white/10 bg-slate-950/50 p-3"
            onScroll={(event) => {
              const container = event.currentTarget;
              const distanceFromBottom =
                container.scrollHeight -
                container.scrollTop -
                container.clientHeight;
              shouldFollowLatestRef.current = distanceFromBottom < 80;
            }}
            ref={chatContainerRef}
          >
            {messages.length === 0 && pendingPrompt === null ? (
              <p className="py-8 text-center text-sm text-slate-500">
                Ask a question to start this conversation.
              </p>
            ) : (
              messages.map((message, index) => (
                <ChatBubble
                  key={message.id}
                  message={message}
                  role={index % 2 === 0 ? "user" : "assistant"}
                />
              ))
            )}
            {pendingPrompt !== null && (
              <div className="flex flex-col items-end gap-2">
                <ChatBubble
                  message={{
                    id: "pending-user-message",
                    text: pendingPrompt,
                    createdDateTime: null,
                    references: [],
                  }}
                  role="user"
                />
                <p
                  className="flex items-center gap-2 px-2 text-xs text-slate-400"
                  role="status"
                >
                  <span className="h-1.5 w-1.5 animate-pulse rounded-full bg-indigo-300" />
                  Work IQ is responding…
                </p>
              </div>
            )}
          </div>
          <form className="mt-4" onSubmit={sendPrompt}>
            <label
              className="text-sm font-medium text-slate-200"
              htmlFor="workiq-prompt"
            >
              Ask Work IQ
            </label>
            <textarea
              className="mt-2 min-h-28 w-full resize-y rounded-xl border border-white/15 bg-slate-950/70 px-3 py-2.5 text-sm leading-6 text-white placeholder:text-slate-500 focus:border-indigo-300 focus:outline-none focus:ring-2 focus:ring-indigo-300/20 disabled:cursor-not-allowed disabled:opacity-60"
              disabled={isSending}
              id="workiq-prompt"
              onChange={(event) => setPrompt(event.target.value)}
              placeholder="Ask a question about your work…"
              value={prompt}
            />
            <button
              className="mt-3 inline-flex min-h-11 w-full items-center justify-center rounded-xl bg-indigo-400 px-4 py-2.5 text-sm font-semibold text-slate-950 transition hover:bg-indigo-300 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-indigo-200 disabled:cursor-not-allowed disabled:opacity-60"
              disabled={isSending || prompt.trim().length === 0}
              type="submit"
            >
              {isSending ? "Sending…" : "Send prompt"}
            </button>
          </form>
        </div>
      ) : (
        <>
          <p className="mt-2 text-sm leading-6 text-slate-300">
            You don’t have a saved Work IQ chat yet. Create one to get started.
          </p>
          <button
            className="mt-4 inline-flex min-h-11 w-full items-center justify-center rounded-xl bg-indigo-400 px-4 py-2.5 text-sm font-semibold text-slate-950 transition hover:bg-indigo-300 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-indigo-200 disabled:cursor-not-allowed disabled:opacity-60"
            disabled={isCreating}
            onClick={createChat}
            type="button"
          >
            {isCreating ? "Creating chat…" : "Create a Work IQ chat"}
          </button>
        </>
      )}

      {errorMessage && (
        <p
          aria-live="polite"
          className="mt-3 rounded-xl border border-rose-300/20 bg-rose-300/5 p-3 text-sm leading-5 text-rose-200"
          role="alert"
        >
          {errorMessage}
        </p>
      )}
    </section>
  );
}

type ChatReference = {
  key: string;
  targetLink: string;
  isCitedInResponse: boolean;
};

type ChatMessage = {
  id: string;
  text: string;
  createdDateTime: string | null;
  references: ChatReference[];
};

function ChatBubble({
  message,
  role,
}: {
  message: ChatMessage;
  role: "user" | "assistant";
}) {
  const citedReferences = message.references.filter(
    (reference) => reference.isCitedInResponse,
  );

  return (
    <article
      className={`flex ${role === "user" ? "justify-end" : "justify-start"}`}
    >
      <div
        className={`max-w-[90%] rounded-2xl px-4 py-3 ${
          role === "user"
            ? "rounded-br-md bg-indigo-400/20 text-indigo-50 ring-1 ring-inset ring-indigo-300/15"
            : "rounded-bl-md bg-white/[0.06] text-slate-200 ring-1 ring-inset ring-white/10"
        }`}
      >
        <p className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">
          {role === "user" ? "You" : "Work IQ"}
        </p>
        <div className="space-y-2 text-sm leading-6">
          {renderMessageText(message.text)}
        </div>
        {citedReferences.length > 0 && (
          <div className="mt-3 border-t border-white/10 pt-3">
            <p className="text-xs font-medium text-slate-400">Sources</p>
            <ul className="mt-2 flex flex-wrap gap-2">
              {citedReferences.map((reference) => (
                <li key={reference.key}>
                  <a
                    className="inline-flex rounded-lg bg-white/5 px-2.5 py-1.5 text-xs text-indigo-200 ring-1 ring-inset ring-white/10 transition hover:bg-white/10"
                    href={reference.targetLink}
                    rel="noreferrer"
                    target="_blank"
                  >
                    {getReferenceLabel(reference.targetLink, reference.key)}
                  </a>
                </li>
              ))}
            </ul>
          </div>
        )}
        {message.createdDateTime && (
          <p className="mt-2 text-right text-[0.6875rem] text-slate-500">
            {formatMessageTime(message.createdDateTime)}
          </p>
        )}
      </div>
    </article>
  );
}

function renderMessageText(text: string): React.ReactNode[] {
  return text.split(/\n{2,}/).map((block, blockIndex) => {
    const lines = block.split("\n");
    const firstHeading = lines[0].match(/^(#{1,6})\s+(.+)$/);

    if (firstHeading && lines.length === 1) {
      return (
        <h4 className="font-semibold text-slate-100" key={blockIndex}>
          {renderInlineMarkdown(firstHeading[2])}
        </h4>
      );
    }

    if (lines.every((line) => /^[-*]\s+/.test(line))) {
      return (
        <ul className="list-disc space-y-1 pl-5" key={blockIndex}>
          {lines.map((line, lineIndex) => (
            <li key={lineIndex}>{renderInlineMarkdown(line.replace(/^[-*]\s+/, ""))}</li>
          ))}
        </ul>
      );
    }

    return (
      <p className="whitespace-pre-wrap break-words" key={blockIndex}>
        {lines.map((line, lineIndex) => (
          <span key={lineIndex}>
            {renderInlineMarkdown(line)}
            {lineIndex < lines.length - 1 && <br />}
          </span>
        ))}
      </p>
    );
  });
}

function renderInlineMarkdown(text: string): React.ReactNode[] {
  const tokenPattern = /(\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)|\*\*([^*]+)\*\*|`([^`]+)`)/g;
  const parts: React.ReactNode[] = [];
  let lastIndex = 0;

  for (const match of text.matchAll(tokenPattern)) {
    const index = match.index ?? 0;
    if (index > lastIndex) parts.push(text.slice(lastIndex, index));

    if (match[2] && match[3]) {
      parts.push(
        <a
          className="text-indigo-200 underline decoration-indigo-300/50 underline-offset-2 hover:text-indigo-100"
          href={match[3]}
          key={`${index}-link`}
          rel="noreferrer"
          target="_blank"
        >
          {match[2]}
        </a>,
      );
    } else if (match[4]) {
      parts.push(<strong key={`${index}-strong`}>{match[4]}</strong>);
    } else if (match[5]) {
      parts.push(
        <code
          className="rounded bg-black/25 px-1 py-0.5 font-mono text-[0.8125rem]"
          key={`${index}-code`}
        >
          {match[5]}
        </code>,
      );
    }

    lastIndex = index + match[0].length;
  }

  if (lastIndex < text.length) parts.push(text.slice(lastIndex));
  return parts;
}

async function readConversation(
  response: Response,
): Promise<{ id: string }> {
  const result: unknown = await response.json();

  if (
    typeof result !== "object" ||
    result === null ||
    !("id" in result) ||
    typeof result.id !== "string" ||
    result.id.trim().length === 0
  ) {
    throw new Error("Work IQ returned a response without a conversation ID.");
  }

  return { id: result.id };
}

async function readChatConversation(
  response: Response,
): Promise<{ messages: ChatMessage[] }> {
  const result: unknown = await response.json();

  if (
    typeof result !== "object" ||
    result === null ||
    !("messages" in result) ||
    !Array.isArray(result.messages)
  ) {
    throw new Error("Work IQ returned a response without a message list.");
  }

  return {
    messages: result.messages.map((message, index) => {
      if (
        typeof message !== "object" ||
        message === null ||
        !("text" in message) ||
        typeof message.text !== "string"
      ) {
        throw new Error(`Work IQ returned an invalid message at position ${index + 1}.`);
      }

      const record = message as Record<string, unknown>;
      return {
        id:
          typeof record.id === "string" && record.id.length > 0
            ? record.id
            : `message-${index}`,
        text: message.text,
        createdDateTime:
          typeof record.createdDateTime === "string"
            ? record.createdDateTime
            : null,
        references: readReferences(record.references),
      };
    }),
  };
}

function readReferences(value: unknown): ChatReference[] {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return [];
  }

  return Object.entries(value).flatMap(([key, reference]) => {
    if (
      typeof reference !== "object" ||
      reference === null ||
      !("targetLink" in reference) ||
      typeof reference.targetLink !== "string" ||
      !isSafeHttpUrl(reference.targetLink)
    ) {
      return [];
    }

    return [
      {
        key,
        targetLink: reference.targetLink,
        isCitedInResponse:
          "isCitedInResponse" in reference &&
          reference.isCitedInResponse === true,
      },
    ];
  });
}

function getReferenceLabel(targetLink: string, fallback: string): string {
  try {
    const url = new URL(targetLink);
    return url.searchParams.get("q") ?? url.hostname ?? fallback;
  } catch {
    return fallback;
  }
}

function formatMessageTime(value: string): string {
  const date = new Date(value);
  return Number.isNaN(date.getTime())
    ? ""
    : date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
}

function isSafeHttpUrl(value: string): boolean {
  try {
    const protocol = new URL(value).protocol;
    return protocol === "http:" || protocol === "https:";
  } catch {
    return false;
  }
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : "An unexpected error occurred.";
}

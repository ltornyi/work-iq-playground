import {
  InteractionRequiredAuthError,
  type AccountInfo,
} from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { useEffect, useState } from "react";

const WORK_IQ_SCOPE =
  "api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask";
const CONVERSATIONS_URL =
  "https://workiq.svc.cloud.microsoft/rest/beta/conversations";

export function WorkIqChatPanel({ account }: { account: AccountInfo }) {
  const { instance } = useMsal();
  const [conversationId, setConversationId] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isCreating, setIsCreating] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const storageKey = `workiq.conversation.${account.homeAccountId}`;

  useEffect(() => {
    setIsLoading(true);
    setConversationId(null);
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

  const createChat = async () => {
    setIsCreating(true);
    setErrorMessage(null);

    try {
      let accessToken: string;
      const tokenRequest = {
        account,
        scopes: [WORK_IQ_SCOPE],
      };

      try {
        ({ accessToken } = await instance.acquireTokenSilent(tokenRequest));
      } catch (error) {
        if (!(error instanceof InteractionRequiredAuthError)) throw error;

        ({ accessToken } = await instance.acquireTokenPopup(tokenRequest));
      }

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

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : "An unexpected error occurred.";
}

import { graphConfig } from "./authConfig";

async function callGraph(endpoint, accessToken, options = {}) {
  const response = await fetch(endpoint, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    ...options,
  });

  if (!response.ok) {
    throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
  }

  if (response.status === 204) return null;
  return response.json();
}

export async function getMe(accessToken) {
  return callGraph(graphConfig.graphMeEndpoint, accessToken);
}

export async function getTodoLists(accessToken) {
  const data = await callGraph(graphConfig.graphTodoEndpoint, accessToken);
  return data.value;
}

export async function getTasksForList(accessToken, listId) {
  const data = await callGraph(
    `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks`,
    accessToken
  );
  return data.value;
}

export async function createTask(accessToken, listId, title) {
  return callGraph(
    `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks`,
    accessToken,
    {
      method: "POST",
      body: JSON.stringify({ title }),
    }
  );
}

export async function updateTask(accessToken, listId, taskId, updates) {
  return callGraph(
    `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks/${taskId}`,
    accessToken,
    {
      method: "PATCH",
      body: JSON.stringify(updates),
    }
  );
}

export async function deleteTask(accessToken, listId, taskId) {
  return callGraph(
    `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks/${taskId}`,
    accessToken,
    { method: "DELETE" }
  );
}

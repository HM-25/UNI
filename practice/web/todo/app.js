// To-do list in plain JavaScript
// Topics: DOM manipulation, events, event delegation, array methods, localStorage

const STORAGE_KEY = "todos";

const form = document.getElementById("todo-form");
const input = document.getElementById("todo-input");
const list = document.getElementById("todo-list");
const counter = document.getElementById("counter");
const clearDone = document.getElementById("clear-done");
const filterButtons = document.querySelectorAll(".filters button");

let todos = loadTodos();
let filter = "all";

function loadTodos() {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
    } catch (e) {
        return [];
    }
}

function saveTodos() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(todos));
    } catch (e) {
        // storage not available, the list still works for this session
    }
}

function visibleTodos() {
    if (filter === "active") return todos.filter(t => !t.done);
    if (filter === "done") return todos.filter(t => t.done);
    return todos;
}

function render() {
    list.innerHTML = "";

    visibleTodos().forEach(todo => {
        const li = document.createElement("li");
        li.dataset.id = todo.id;
        if (todo.done) li.classList.add("done");

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = todo.done;
        checkbox.className = "toggle";

        const text = document.createElement("span");
        text.textContent = todo.text;   // textContent, not innerHTML, so user input is never run as HTML

        const del = document.createElement("button");
        del.className = "delete";
        del.textContent = "✕";
        del.title = "Delete";

        li.append(checkbox, text, del);
        list.appendChild(li);
    });

    const left = todos.filter(t => !t.done).length;
    counter.textContent = `${left} item${left === 1 ? "" : "s"} left`;
    saveTodos();
}

form.addEventListener("submit", event => {
    event.preventDefault();
    const text = input.value.trim();
    if (!text) return;

    todos.push({ id: Date.now(), text: text, done: false });
    input.value = "";
    render();
});

// One listener on the list instead of one per item (event delegation)
list.addEventListener("click", event => {
    const li = event.target.closest("li");
    if (!li) return;
    const id = Number(li.dataset.id);

    if (event.target.classList.contains("toggle")) {
        const todo = todos.find(t => t.id === id);
        todo.done = !todo.done;
    } else if (event.target.classList.contains("delete")) {
        todos = todos.filter(t => t.id !== id);
    } else {
        return;
    }
    render();
});

filterButtons.forEach(button => {
    button.addEventListener("click", () => {
        filter = button.dataset.filter;
        filterButtons.forEach(b => b.classList.toggle("active", b === button));
        render();
    });
});

clearDone.addEventListener("click", () => {
    todos = todos.filter(t => !t.done);
    render();
});

render();

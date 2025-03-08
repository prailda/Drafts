// Добавляем функцию переключения видов
function switchView(viewName) {
    state.views.current = viewName;
    render();
}

// Основная функция рендеринга текущего вида
function renderCurrentView() {
    switch (state.views.current) {
        case 'kanban':
            return renderKanbanBoard();
        case 'timeline':
            return renderTimeline();
        case 'stats':
            return renderStatistics();
        case 'tree':
        default:
            return renderTaskTree();
    }
}

// Рендеринг дерева задач
function renderTaskTree() {
    const tasks = filterTasks(state.tasks.filter(t => 
        t.projectId === state.currentProject?.id && !t.parentId
    ));

    if (!tasks.length) {
        return '<p>Нет задач</p>';
    }

    return `
        <div class="filter-panel">
            ${renderFilters()}
        </div>
        <div class="task-list">
            ${tasks.map(task => renderTaskItem(task)).join('')}
        </div>
    `;
}

// Рендеринг одной задачи
function renderTaskItem(task) {
    return `
        <div class="task-item">
            <div class="task-header">
                <div class="task-header-content">
                    <h3 class="task-title">${task.name}</h3>
                    <div class="task-badges">
                        <span class="badge badge-${getPriorityClass(task.priority)}">
                            ${getPriorityLabel(task.priority)}
                        </span>
                        <span class="badge badge-${getStatusClass(task.status)}">
                            ${getStatusLabel(task.status)}
                        </span>
                        ${task.categoryId ? `
                            <span class="badge" style="background-color: ${getCategoryColor(task.categoryId)}">
                                ${getCategoryName(task.categoryId)}
                            </span>
                        ` : ''}
                    </div>
                </div>
                <div class="task-actions">
                    <button class="btn btn-small" onclick="toggleTaskCompletion('${task.id}')">
                        ${task.status === 'completed' ? '✓' : '⬜'}
                    </button>
                    <button class="btn btn-small" onclick="showEditTaskModal('${task.id}')">
                        ✎
                    </button>
                    <button class="btn btn-small" onclick="deleteTask('${task.id}')">
                        ×
                    </button>
                </div>
            </div>
            
            <div class="task-content">
                ${task.description ? `<p>${task.description}</p>` : ''}
                
                <div class="task-dates">
                    ${task.startDate ? `
                        <div>Начало: ${new Date(task.startDate).toLocaleDateString()}</div>
                    ` : ''}
                    ${task.dueDate ? `
                        <div>Срок: ${new Date(task.dueDate).toLocaleDateString()}</div>
                    ` : ''}
                </div>

                ${renderTaskAttachments(task)}
                ${renderTaskChecklist(task)}
                ${renderTaskDependencies(task)}
                ${renderTaskComments(task)}
            </div>

            <div class="task-children">
                ${renderSubtasks(task.id)}
            </div>
        </div>
    `;
}

// Рендеринг канбан-доски
function renderKanbanBoard() {
    const statuses = ['new', 'inProgress', 'onHold', 'completed'];
    const tasks = filterTasks(state.tasks.filter(t => t.projectId === state.currentProject?.id));

    return `
        <div class="filter-panel">
            ${renderFilters()}
        </div>
        <div class="kanban-board">
            ${statuses.map(status => `
                <div class="kanban-column">
                    <div class="kanban-column-header">
                        <h3>${getStatusLabel(status)}</h3>
                        <span class="task-count">
                            ${tasks.filter(t => t.status === status).length}
                        </span>
                    </div>
                    <div class="kanban-column-content" 
                         ondrop="dropTask(event)" 
                         ondragover="allowDrop(event)" 
                         data-status="${status}">
                        ${tasks.filter(t => t.status === status)
                            .map(task => renderKanbanCard(task))
                            .join('')}
                    </div>
                </div>
            `).join('')}
        </div>
    `;
}

// Рендеринг карточки для канбан-доски
function renderKanbanCard(task) {
    return `
        <div class="kanban-card" 
             draggable="true" 
             ondragstart="dragTask(event)" 
             data-task-id="${task.id}">
            <div class="kanban-card-header">
                <h4>${task.name}</h4>
                <div class="task-badges">
                    <span class="badge badge-${getPriorityClass(task.priority)}">
                        ${getPriorityLabel(task.priority)}
                    </span>
                    ${task.categoryId ? `
                        <span class="badge" style="background-color: ${getCategoryColor(task.categoryId)}">
                            ${getCategoryName(task.categoryId)}
                        </span>
                    ` : ''}
                </div>
            </div>
            ${task.dueDate ? `
                <div class="kanban-card-date">
                    Срок: ${new Date(task.dueDate).toLocaleDateString()}
                </div>
            ` : ''}
            <div class="progress">
                <div class="progress-bar" style="width: ${getTaskProgress(task)}%"></div>
            </div>
        </div>
    `;
}

// Функции для Drag & Drop
function dragTask(event) {
    event.dataTransfer.setData('taskId', event.target.dataset.taskId);
}

function allowDrop(event) {
    event.preventDefault();
}

function dropTask(event) {
    event.preventDefault();
    const taskId = event.dataTransfer.getData('taskId');
    const newStatus = event.target.closest('.kanban-column-content').dataset.status;
    
    const task = state.tasks.find(t => t.id === taskId);
    if (task && task.status !== newStatus) {
        const oldStatus = task.status;
        task.status = newStatus;
        task.updatedAt = new Date().toISOString();

        // Добавляем запись в историю
        task.history = task.history || [];
        task.history.push({
            id: generateId(),
            type: 'status_change',
            oldValue: oldStatus,
            newValue: newStatus,
            timestamp: new Date().toISOString()
        });

        saveState();
        render();
    }
}

// Компоненты фильтрации
function renderFilters() {
    return `
        <div class="filters">
            <div class="filter-group">
                <label>Категория</label>
                <select onchange="updateFilter('category', this.value)">
                    <option value="">Все категории</option>
                    ${state.categories
                        .filter(c => c.projectId === state.currentProject?.id)
                        .map(category => `
                            <option value="${category.id}" 
                                    ${state.views.filters.category === category.id ? 'selected' : ''}>
                                ${category.name}
                            </option>
                        `).join('')}
                </select>
            </div>
            <div class="filter-group">
                <label>Приоритет</label>
                <select onchange="updateFilter('priority', this.value)">
                    <option value="">Все приоритеты</option>
                    <option value="high" ${state.views.filters.priority === 'high' ? 'selected' : ''}>
                        Высокий
                    </option>
                    <option value="medium" ${state.views.filters.priority === 'medium' ? 'selected' : ''}>
                        Средний
                    </option>
                    <option value="low" ${state.views.filters.priority === 'low' ? 'selected' : ''}>
                        Низкий
                    </option>
                </select>
            </div>
            <div class="filter-group">
                <label>Статус</label>
                <select onchange="updateFilter('status', this.value)">
                    <option value="">Все статусы</option>
                    <option value="new" ${state.views.filters.status === 'new' ? 'selected' : ''}>
                        Новые
                    </option>
                    <option value="inProgress" ${state.views.filters.status === 'inProgress' ? 'selected' : ''}>
                        В работе
                    </option>
                    <option value="onHold" ${state.views.filters.status === 'onHold' ? 'selected' : ''}>
                        На паузе
                    </option>
                    <option value="completed" ${state.views.filters.status === 'completed' ? 'selected' : ''}>
                        Завершенные
                    </option>
                </select>
            </div>
            <div class="filter-group">
                <label>Поиск</label>
                <input type="text" 
                       placeholder="Поиск по названию..." 
                       value="${state.views.filters.search}"
                       oninput="updateFilter('search', this.value)">
            </div>
        </div>
    `;
}

function updateFilter(type, value) {
    state.views.filters[type] = value;
    render();
}

function filterTasks(tasks) {
    const filters = state.views.filters;
    return tasks.filter(task => {
        if (filters.category && task.categoryId !== filters.category) return false;
        if (filters.priority && task.priority !== filters.priority) return false;
        if (filters.status && task.status !== filters.status) return false;
        if (filters.search && !task.name.toLowerCase().includes(filters.search.toLowerCase())) return false;
        return true;
    });
}

// Вспомогательные функции для отображения данных задачи
function renderTaskAttachments(task) {
    if (!task.attachments?.length) return '';
    
    return `
        <div class="task-attachments">
            <h4>Вложения</h4>
            <div class="attachments-list">
                ${task.attachments.map(file => `
                    <div class="attachment-item">
                        <a href="${file.url}" target="_blank" download="${file.name}">
                            ${file.name}
                        </a>
                        <span class="attachment-date">
                            ${new Date(file.createdAt).toLocaleDateString()}
                        </span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

function renderTaskComments(task) {
    if (!task.comments?.length) return '';
    
    return `
        <div class="task-comments">
            <h4>Комментарии</h4>
            <div class="comments-list">
                ${task.comments.map(comment => `
                    <div class="comment-item">
                        <div class="comment-header">
                            <span class="comment-date">
                                ${new Date(comment.createdAt).toLocaleString()}
                            </span>
                        </div>
                        <div class="comment-content">
                            ${comment.text}
                        </div>
                    </div>
                `).join('')}
            </div>
            <div class="comment-form">
                <textarea placeholder="Добавить комментарий..." 
                          onkeydown="if(event.key === 'Enter' && !event.shiftKey) { 
                              event.preventDefault(); 
                              addComment('${task.id}', this.value); 
                              this.value = ''; 
                          }"></textarea>
            </div>
        </div>
    `;
}

// Функции для работы с комментариями
function addComment(taskId, text) {
    if (!text.trim()) return;
    
    const task = state.tasks.find(t => t.id === taskId);
    if (task) {
        task.comments = task.comments || [];
        task.comments.push({
            id: generateId(),
            text: text.trim(),
            createdAt: new Date().toISOString()
        });
        
        saveState();
        render();
    }
}

// Валидация при отправке формы задачи
function validateTaskForm(form) {
    const startDate = form.taskStartDate.value;
    const dueDate = form.taskDueDate.value;
    const timeEstimate = parseFloat(form.taskTimeEstimate.value);

    if (startDate && dueDate && new Date(dueDate) < new Date(startDate)) {
        alert('Срок выполнения не может быть раньше даты начала');
        return false;
    }

    if (timeEstimate < 0) {
        alert('Оценка времени не может быть отрицательной');
        return false;
    }

    return true;
}

// Расширяем handleTaskSubmit для валидации
const originalHandleTaskSubmit = handleTaskSubmit;
handleTaskSubmit = function(event) {
    if (!validateTaskForm(event.target)) {
        event.preventDefault();
        return;
    }
    originalHandleTaskSubmit.call(this, event);
};

// Функция для расчета прогресса задачи
function getTaskProgress(task) {
    if (task.status === 'completed') return 100;
    if (!task.checklist?.length) return 0;
    
    const completed = task.checklist.filter(item => item.completed).length;
    return Math.round((completed / task.checklist.length) * 100);
}

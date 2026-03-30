// ===== Section definitions =====
const SECTIONS = [
    {
        id: 'sec1',
        title: '一、目的',
        hint: `明确制度制定的核心目标（需量化可评估，避免模糊表述）。\n\n示例：\n为规范公司合同管理行为，降低法律风险，确保合同履约率≥98%，特制定本制度。`
    },
    {
        id: 'sec2',
        title: '二、适用范围',
        hint: `适用对象（如全公司/特定部门/岗位）；\n适用业务场景（如采购类合同/劳动合同）；\n排除范围（如"本制度不适用于XX子公司"）。`
    },
    {
        id: 'sec3',
        title: '三、规范性引用文件',
        hint: `（可选，没有则填无）`
    },
    {
        id: 'sec4',
        title: '四、术语和定义',
        hint: `（可选，没有则填无）\n对制度中关键术语进行解释（避免歧义）。\n\n示例：\n"重大合同：指单笔金额≥100万元或涉及公司核心技术的合同。"\n"无。"`
    },
    {
        id: 'sec5',
        title: '五、基本原则',
        hint: `（可选，没有则填无）\n列出3-5条核心管理原则（体现企业价值观或方法论）。\n\n示例：\n1．合法性原则：所有合同条款必须符合《民法典》规定；\n2．风险优先原则：金额每增加50万元，审批层级提高一级...`
    },
    {
        id: 'sec6',
        title: '六、职责',
        hint: `具体职责：合同合法性审查；\n权限边界：否决权（仅限法律条款）；`
    },
    {
        id: 'sec7',
        title: '七、管理要求',
        hint: `7.1核心规则\n分条款列出禁止性、强制性和指引性要求。\n\n示例：\n禁止任何部门未经招标签订≥20万元的采购合同；\n所有合同必须使用公司标准模板，修改条款需法务部备案...`
    },
    {
        id: 'sec8',
        title: '八、特殊处理机制',
        hint: `（可选，没有则填无）\n规定特殊情况的处理方式（需明确申请路径）。\n\n示例：\n"因紧急生产需要可先执行后补签合同，但需在3个工作日内提交CEO特批说明。"`
    },
    {
        id: 'sec9',
        title: '九、监管与问责',
        hint: `8.1监督\n定期检查（如"每季度审计部抽查20%合同"）；\n触发检查（如"单笔合同违约金额≥10万元启动专项审计"）。\n\n8.2惩罚（重大制度必填，一般制度若没有则填无）\n违规行为与处罚措施对应（需引用《员工奖惩制度》具体条款）。\n\n示例：私自修改合同条款者，按造成损失的200%追偿；未按时归档合同，扣减责任人当月绩效10%...\n\n重大制度，一般制度区分详见《制度新增&修订&废除流程说明书》4.3一般制度与重大制度的区分原则`
    },
    {
        id: 'sec10',
        title: '十、附则',
        hint: `XXXXXX。\n本办法自会签发布之日起生效，新制度需设置试行期（重大制度通常3-6个月，一般制度1个月），并根据反馈调整。\n本办法的最终解释权归XXX。`
    },
    {
        id: 'sec11',
        title: '十一、附录',
        hint: `（可选，没有则填无）\n1．《XX审批流程图》（必附）；\n2．《XX标准模板》（如合同/表单）；\n3．《XX记录表》（如检查台账）。`
    }
];

// ===== Render all sections =====
function renderSections() {
    const container = document.getElementById('editorContainer');
    SECTIONS.forEach(sec => {
        const card = document.createElement('div');
        card.className = 'section-card';
        card.innerHTML = `
            <div class="section-header">
                <h3 class="section-title">${sec.title}</h3>
                <div class="hint-icon">?<div class="hint-popup">${sec.hint}</div></div>
            </div>
            <div class="editor-wrapper">
                <div id="${sec.id}" class="editor-inline"></div>
            </div>
            <div class="action-row">
                <button class="btn-check-ai" id="btn-${sec.id}" onclick="aiCheck('${sec.id}', '${sec.title}')">
                    完成
                </button>
                <span class="ai-disclaimer">AI 审核结果仅供参考，请以实际需求为准</span>
            </div>
            <div class="ai-result" id="result-${sec.id}"></div>
        `;
        container.appendChild(card);
    });
}

// ===== Initialize TinyMCE (inline mode) =====
function initEditors() {
    tinymce.init({
        selector: SECTIONS.map(s => '#' + s.id).join(','),
        inline: true,
        fixed_toolbar_container: '#tinymce-toolbar-container',
        language: 'zh_CN',
        language_url: '/static/tinymce/langs/zh_CN.js',
        skin_url: '/tinymce/skins/ui/oxide',
        content_css: '/tinymce/skins/content/default/content.min.css',
        menubar: 'file edit view insert format table',
        plugins: [
            'advlist', 'autolink', 'lists', 'link', 'image', 'charmap', 'preview',
            'searchreplace', 'visualblocks',
            'insertdatetime', 'table', 'wordcount'
        ],
        toolbar: 'undo redo | blocks | bold italic underline strikethrough | ' +
                 'forecolor backcolor | alignleft aligncenter alignright alignjustify firstindent | ' +
                 'bullist numlist outdent indent | image drawio table cellvalign | removeformat',
        formats: {
            firstindent: { selector: 'p,div', styles: { 'text-indent': '2em' } }
        },
        setup: function(editor) {
            editor.ui.registry.addToggleButton('firstindent', {
                text: '首行缩进',
                tooltip: '首行缩进 2字符',
                onAction: function() {
                    editor.formatter.toggle('firstindent');
                    editor.nodeChanged();
                },
                onSetup: function(api) {
                    editor.on('NodeChange', function() {
                        api.setActive(editor.formatter.match('firstindent'));
                    });
                }
            });

            // Vertical alignment menu button for table cells
            editor.ui.registry.addMenuButton('cellvalign', {
                text: '垂直对齐',
                tooltip: '单元格垂直对齐',
                fetch: function(callback) {
                    callback([
                        {
                            type: 'menuitem',
                            text: '顶部对齐',
                            onAction: function() {
                                var selectedCells = editor.dom.select('td[data-mce-selected],th[data-mce-selected]');
                                if (selectedCells.length > 0) {
                                    selectedCells.forEach(function(cell) {
                                        editor.dom.setStyle(cell, 'vertical-align', 'top');
                                    });
                                } else {
                                    var cell = editor.dom.getParent(editor.selection.getStart(), 'td,th');
                                    if (cell) {
                                        editor.dom.setStyle(cell, 'vertical-align', 'top');
                                    }
                                }
                                editor.nodeChanged();
                            }
                        },
                        {
                            type: 'menuitem',
                            text: '垂直居中',
                            onAction: function() {
                                var selectedCells = editor.dom.select('td[data-mce-selected],th[data-mce-selected]');
                                if (selectedCells.length > 0) {
                                    selectedCells.forEach(function(cell) {
                                        editor.dom.setStyle(cell, 'vertical-align', 'middle');
                                    });
                                } else {
                                    var cell = editor.dom.getParent(editor.selection.getStart(), 'td,th');
                                    if (cell) {
                                        editor.dom.setStyle(cell, 'vertical-align', 'middle');
                                    }
                                }
                                editor.nodeChanged();
                            }
                        },
                        {
                            type: 'menuitem',
                            text: '底部对齐',
                            onAction: function() {
                                var selectedCells = editor.dom.select('td[data-mce-selected],th[data-mce-selected]');
                                if (selectedCells.length > 0) {
                                    selectedCells.forEach(function(cell) {
                                        editor.dom.setStyle(cell, 'vertical-align', 'bottom');
                                    });
                                } else {
                                    var cell = editor.dom.getParent(editor.selection.getStart(), 'td,th');
                                    if (cell) {
                                        editor.dom.setStyle(cell, 'vertical-align', 'bottom');
                                    }
                                }
                                editor.nodeChanged();
                            }
                        }
                    ]);
                }
            });

            // Draw.io diagram button
            editor.ui.registry.addButton('drawio', {
                text: '画图',
                tooltip: '使用 Draw.io 绘制架构图 / 流程图',
                onAction: function() {
                    var node = editor.selection.getNode();
                    if (node.nodeName === 'IMG' && node.hasAttribute('data-drawio-xml')) {
                        _openDrawio(editor, node);
                    } else {
                        _openDrawio(editor, null);
                    }
                }
            });

            // Double-click to edit existing Draw.io diagrams
            editor.on('dblclick', function(e) {
                if (e.target.nodeName === 'IMG' && e.target.hasAttribute('data-drawio-xml')) {
                    e.preventDefault();
                    _openDrawio(editor, e.target);
                }
            });

            // Track editor focus for toolbar switching and visibility
            editor.on('focus', function() {
                console.log('[DEBUG] editor focus, secId:', editor.id);
                _onEditorFocus(editor);
                // Always show sticky toolbar when editing, even at page top
                _activateToolbarForced();
            });
            editor.on('blur', function() {
                console.log('[DEBUG] editor blur');
                // Restore normal scroll-based visibility after blur
                _onEditorBlur();
            });
        },
        extended_valid_elements: 'img[class|src|border|alt|title|width|height|style|data-drawio-xml|loading]',
        table_default_styles: {
            'border-collapse': 'collapse',
            'width': '100%'
        },
        table_grid: false,
        table_cell_advtab: true,
        images_upload_url: '/api/upload-image',
        automatic_uploads: true,
        file_picker_types: 'image',
        paste_preprocess: function(plugin, args) {
            var c = args.content;
            c = c.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, '');
            c = c.replace(/\[if\s+!support\w+\]/gi, '');
            c = c.replace(/\[endif\]/gi, '');
            c = c.replace(/<\?xml[\s\S]*?\?>/gi, '');
            c = c.replace(/<style[\s\S]*?<\/style>/gi, '');
            c = c.replace(/<\/?\w+:[^>]*>/gi, '');
            c = c.replace(/\s*xmlns:\w+="[^"]*"/gi, '');
            c = c.replace(/\s*class="Mso[^"]*"/gi, '');
            c = c.replace(/mso-[a-z\-]+\s*:[^;"]*;?/gi, '');
            c = c.replace(/\s*style="\s*"/gi, '');
            c = c.replace(/<\/?font[^>]*>/gi, '');
            c = c.replace(/<span[^>]*mso-spacerun[^>]*>[\s\u00a0]*<\/span>/gi, '');
            c = c.replace(/<span><\/span>/gi, '');
            c = c.replace(/\s*(xml:)?lang="[^"]*"/gi, '');
            args.content = c;
        },
        branding: false,
        promotion: false,
        license_key: 'gpl',
        convert_urls: false,
        relative_urls: false,
        init_instance_callback: function() {
            _editorsReady++;
            if (_editorsReady === SECTIONS.length) {
                _initDrafts();
                _startAutoSave();
            }
        }
    });
}

// ===== Multi-Document Storage & Sidebar =====
var DRAFTS_KEY = 'zhidu_drafts';
var AUTOSAVE_INTERVAL = 30000;
var _editorsReady = 0;
var _autosaveTimer = null;
var _draftsData = null; // { version, activeId, docs }

function _generateId() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function _loadDraftsData() {
    // Try new format first
    try {
        var raw = localStorage.getItem(DRAFTS_KEY);
        if (raw) {
            var data = JSON.parse(raw);
            if (data && data.version === 2 && data.docs) {
                _draftsData = data;
                return;
            }
        }
    } catch (e) { /* corrupt */ }

    // Migrate from old single-draft format
    try {
        var oldRaw = localStorage.getItem('zhidu_draft');
        if (oldRaw) {
            var oldDraft = JSON.parse(oldRaw);
            if (oldDraft && oldDraft.timestamp) {
                var id = _generateId();
                _draftsData = {
                    version: 2,
                    activeId: id,
                    docs: {}
                };
                _draftsData.docs[id] = oldDraft;
                _saveDraftsData();
                localStorage.removeItem('zhidu_draft');
                return;
            }
        }
    } catch (e) { /* corrupt */ }

    // Fresh start
    var id = _generateId();
    _draftsData = {
        version: 2,
        activeId: id,
        docs: {}
    };
    _draftsData.docs[id] = { docName: '', sections: {}, timestamp: Date.now() };
    _saveDraftsData();
}

function _saveDraftsData() {
    try {
        localStorage.setItem(DRAFTS_KEY, JSON.stringify(_draftsData));
    } catch (e) {
        alert('浏览器存储空间已满，请导出并删除部分旧文档后重试。');
    }
}

function _collectDraft() {
    return {
        docName: document.getElementById('docName').value.trim(),
        sections: (function() {
            var s = {};
            for (var i = 0; i < SECTIONS.length; i++) {
                var editor = tinymce.get(SECTIONS[i].id);
                if (editor) s[SECTIONS[i].id] = editor.getContent();
            }
            return s;
        })(),
        timestamp: Date.now()
    };
}

function _saveCurrentDocToMemory() {
    if (!_draftsData || !_draftsData.activeId) return;
    _draftsData.docs[_draftsData.activeId] = _collectDraft();
}

function _loadDocIntoEditors(docId) {
    var doc = _draftsData.docs[docId];
    if (!doc) return;
    // Set doc name
    var dn1 = document.getElementById('docName');
    var dn2 = document.getElementById('docNameToolbar');
    dn1.value = doc.docName || '';
    if (dn2) dn2.value = doc.docName || '';
    // Set editor contents
    for (var i = 0; i < SECTIONS.length; i++) {
        var secId = SECTIONS[i].id;
        var editor = tinymce.get(secId);
        if (editor) {
            editor.setContent(doc.sections && doc.sections[secId] ? doc.sections[secId] : '');
        }
    }
    // Clear all AI results
    for (var j = 0; j < SECTIONS.length; j++) {
        var resultEl = document.getElementById('result-' + SECTIONS[j].id);
        if (resultEl) resultEl.innerHTML = '';
    }
}

function _clearEditors() {
    document.getElementById('docName').value = '';
    var dn2 = document.getElementById('docNameToolbar');
    if (dn2) dn2.value = '';
    for (var i = 0; i < SECTIONS.length; i++) {
        var editor = tinymce.get(SECTIONS[i].id);
        if (editor) editor.setContent('');
    }
}

function _autoSave() {
    if (!_draftsData) return;
    _saveCurrentDocToMemory();
    _saveDraftsData();
}

function _startAutoSave() {
    _autosaveTimer = setInterval(_autoSave, AUTOSAVE_INTERVAL);
    window.addEventListener('beforeunload', _autoSave);
}

function _initDrafts() {
    _loadDraftsData();
    _loadDocIntoEditors(_draftsData.activeId);
    renderSidebarList();
    // Restore sidebar open state
    if (localStorage.getItem('zhidu_sidebar_open') === 'true') {
        document.body.classList.add('sidebar-open');
    }
}

// ===== Sidebar UI =====
function toggleSidebar() {
    var open = document.body.classList.toggle('sidebar-open');
    localStorage.setItem('zhidu_sidebar_open', open ? 'true' : 'false');
}

function renderSidebarList() {
    var list = document.getElementById('sidebarDocList');
    if (!list || !_draftsData) return;
    list.innerHTML = '';

    // Sort docs by timestamp desc
    var ids = Object.keys(_draftsData.docs);
    ids.sort(function(a, b) {
        return (_draftsData.docs[b].timestamp || 0) - (_draftsData.docs[a].timestamp || 0);
    });

    for (var i = 0; i < ids.length; i++) {
        var id = ids[i];
        var doc = _draftsData.docs[id];
        var isActive = id === _draftsData.activeId;
        var name = doc.docName || '未命名制度';
        var time = doc.timestamp ? new Date(doc.timestamp).toLocaleString('zh-CN') : '';

        var item = document.createElement('div');
        item.className = 'sidebar-doc-item' + (isActive ? ' active' : '');
        item.setAttribute('data-doc-id', id);
        item.innerHTML =
            '<div class="sidebar-doc-name">' + _escHtml(name) + '</div>' +
            '<div class="sidebar-doc-time">' + time + '</div>' +
            '<button class="sidebar-doc-delete" title="删除" data-del-id="' + id + '">&times;</button>';

        // Click to switch (but not on delete button)
        (function(docId) {
            item.addEventListener('click', function(e) {
                if (e.target.closest('.sidebar-doc-delete')) return;
                switchToDoc(docId);
            });
        })(id);

        // Delete button
        (function(docId) {
            item.querySelector('.sidebar-doc-delete').addEventListener('click', function(e) {
                e.stopPropagation();
                deleteDoc(docId);
            });
        })(id);

        list.appendChild(item);
    }
}

function _escHtml(str) {
    var d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
}

function createNewDoc() {
    // Save current doc first
    _saveCurrentDocToMemory();
    // Create new empty doc
    var id = _generateId();
    _draftsData.docs[id] = { docName: '', sections: {}, timestamp: Date.now() };
    _draftsData.activeId = id;
    _saveDraftsData();
    _clearEditors();
    renderSidebarList();
    // Scroll to top
    window.scrollTo(0, 0);
}

function switchToDoc(docId) {
    if (docId === _draftsData.activeId) return;
    // Save current
    _saveCurrentDocToMemory();
    // Switch
    _draftsData.activeId = docId;
    _saveDraftsData();
    _loadDocIntoEditors(docId);
    renderSidebarList();
    window.scrollTo(0, 0);
}

function deleteDoc(docId) {
    var doc = _draftsData.docs[docId];
    var name = (doc && doc.docName) ? '「' + doc.docName + '」' : '未命名制度';
    if (!confirm('确定删除 ' + name + '？此操作不可恢复。')) return;

    delete _draftsData.docs[docId];

    // If deleted the active doc, switch to another or create new
    if (docId === _draftsData.activeId) {
        var remaining = Object.keys(_draftsData.docs);
        if (remaining.length > 0) {
            _draftsData.activeId = remaining[0];
            _loadDocIntoEditors(_draftsData.activeId);
        } else {
            var newId = _generateId();
            _draftsData.docs[newId] = { docName: '', sections: {}, timestamp: Date.now() };
            _draftsData.activeId = newId;
            _clearEditors();
        }
    }

    _saveDraftsData();
    renderSidebarList();
}

function clearDraft() {
    if (!confirm('确定删除当前文档的所有内容？')) return;
    _clearEditors();
    if (_draftsData && _draftsData.activeId) {
        _draftsData.docs[_draftsData.activeId] = { docName: '', sections: {}, timestamp: Date.now() };
        _saveDraftsData();
        renderSidebarList();
    }
}

// ===== Markdown renderer (using marked.js) =====
function renderMarkdown(text) {
    if (!text) return '';
    return marked.parse(text);
}

// ===== AI Check =====
async function aiCheck(secId, title) {
    const editor = tinymce.get(secId);
    if (!editor) return;

    const content = editor.getContent();
    const plainText = editor.getContent({ format: 'text' }).trim();

    if (!plainText) {
        const resultDiv = document.getElementById('result-' + secId);
        resultDiv.textContent = '内容为空，请先输入内容再进行检查。';
        resultDiv.className = 'ai-result show error';
        return;
    }

    const btn = document.getElementById('btn-' + secId);
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner-sm"></span>AI 检查中...';

    const resultDiv = document.getElementById('result-' + secId);
    resultDiv.className = 'ai-result';
    resultDiv.textContent = '';
    resultDiv.removeAttribute('style');

    try {
        const resp = await fetch('/api/ai-check', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ title, content, doc_name: document.getElementById('docName').value.trim() })
        });
        const data = await resp.json();

        if (data.error) {
            resultDiv.textContent = data.error;
            resultDiv.className = 'ai-result show error';
        } else {
            resultDiv.innerHTML = renderMarkdown(data.result);
            resultDiv.className = 'ai-result show';
        }
    } catch (e) {
        resultDiv.textContent = '请求失败：' + e.message;
        resultDiv.className = 'ai-result show error';
    } finally {
        btn.disabled = false;
        btn.innerHTML = '完成';
    }
}

// ===== Export DOCX =====
async function exportDocx() {
    const sections = [];
    for (const sec of SECTIONS) {
        const editor = tinymce.get(sec.id);
        const content = editor ? editor.getContent() : '';
        sections.push({ title: sec.title, content });
    }

    const btn = document.getElementById('exportBtn');
    btn.disabled = true;
    btn.textContent = '导出中...';

    try {
        const resp = await fetch('/api/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ sections, doc_name: document.getElementById('docName').value.trim() })
        });

        if (!resp.ok) throw new Error('导出失败');

        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = (document.getElementById('docName').value.trim() || '制度文档') + '.docx';
        a.click();
        URL.revokeObjectURL(url);
    } catch (e) {
        alert('导出失败: ' + e.message);
    } finally {
        btn.disabled = false;
        btn.textContent = '导出 DOCX';
    }
}

// ===== Draw.io Integration =====
var _drawio = { frame: null, editor: null, img: null, xml: '', exitAfterExport: false };
var _drawioBaseUrl = 'https://embed.diagrams.net';

// Fetch configurable Draw.io URL on startup
(function() {
    fetch('/api/site-config').then(function(r) { return r.json(); }).then(function(cfg) {
        if (cfg.drawio_url) _drawioBaseUrl = cfg.drawio_url.replace(/\/+$/, '');
    }).catch(function() {});
})();

function _openDrawio(editor, existingImg) {
    _drawio.editor = editor;
    _drawio.exitAfterExport = false;

    if (existingImg) {
        _drawio.img = existingImg;
        _drawio.xml = _decodeDrawioXml(existingImg.getAttribute('data-drawio-xml') || '');
    } else {
        _drawio.img = null;
        _drawio.xml = '';
    }

    var overlay = document.createElement('div');
    overlay.id = 'drawio-overlay';
    overlay.className = 'drawio-overlay';

    var frame = document.createElement('iframe');
    frame.className = 'drawio-frame';
    frame.src = _drawioBaseUrl + '/?embed=1&proto=json&spin=1&lang=zh&saveAndExit=1';

    overlay.appendChild(frame);
    document.body.appendChild(overlay);
    _drawio.frame = frame;

    window.addEventListener('message', _onDrawioMsg);
}

function _onDrawioMsg(evt) {
    if (!_drawio.frame || evt.source !== _drawio.frame.contentWindow) return;

    var msg;
    try { msg = JSON.parse(evt.data); } catch (e) { return; }

    if (msg.event === 'init') {
        // Editor ready — load diagram
        _drawio.frame.contentWindow.postMessage(JSON.stringify({
            action: 'load',
            xml: _drawio.xml || '',
            autosave: 0
        }), '*');
    } else if (msg.event === 'save') {
        // User clicked Save or Save&Exit — request PNG export
        _drawio.xml = msg.xml;
        _drawio.exitAfterExport = !!msg.exit;
        _drawio.frame.contentWindow.postMessage(JSON.stringify({
            action: 'export',
            format: 'png',
            xml: msg.xml,
            spin: '导出中...'
        }), '*');
    } else if (msg.event === 'export') {
        // Received PNG — insert into TinyMCE
        _insertDrawioImg(msg.data, _drawio.xml);
        if (_drawio.exitAfterExport) {
            _closeDrawio();
        }
    } else if (msg.event === 'exit') {
        _closeDrawio();
    }
}

function _insertDrawioImg(dataUrl, xml) {
    var editor = _drawio.editor;
    if (!editor) return;

    var xmlB64 = _encodeDrawioXml(xml);

    if (_drawio.img) {
        // Update existing diagram
        editor.dom.setAttribs(_drawio.img, {
            'src': dataUrl,
            'data-drawio-xml': xmlB64
        });
        editor.nodeChanged();
    } else {
        // Insert new diagram
        editor.insertContent(
            '<p><img src="' + dataUrl + '" data-drawio-xml="' + xmlB64 + '" style="max-width:100%;" /></p>'
        );
        // Track for subsequent saves in same session
        var imgs = editor.dom.select('img[data-drawio-xml="' + xmlB64 + '"]');
        if (imgs.length > 0) {
            _drawio.img = imgs[imgs.length - 1];
        }
    }
}

function _closeDrawio() {
    window.removeEventListener('message', _onDrawioMsg);
    var overlay = document.getElementById('drawio-overlay');
    if (overlay) overlay.remove();
    _drawio = { frame: null, editor: null, img: null, xml: '', exitAfterExport: false };
}

function _encodeDrawioXml(xml) {
    try { return btoa(unescape(encodeURIComponent(xml))); }
    catch (e) { return ''; }
}

function _decodeDrawioXml(b64) {
    try { return decodeURIComponent(escape(atob(b64))); }
    catch (e) { return ''; }
}

// ===== Sticky Toolbar & Editor Switching =====
var _activeSecId = null;
var _scrollSwitchTimer = null;
var _toolbarForcedMode = false;  // true when editor is focused (click mode)

function _onEditorFocus(editor) {
    var secId = editor.id;
    if (secId !== _activeSecId) {
        _activeSecId = secId;
        _updateSecLabel(secId);
    }
}

function _updateSecLabel(secId) {
    var label = document.getElementById('toolbarSecLabel');
    if (!label) return;
    for (var i = 0; i < SECTIONS.length; i++) {
        if (SECTIONS[i].id === secId) {
            label.textContent = SECTIONS[i].title;
            return;
        }
    }
}

function _updateHeaderVisibility() {
    var navbar = document.getElementById('mainNavbar');
    var toolbar = document.getElementById('stickyToolbar');
    // When editor is focused (click mode), toolbar is already forced visible - do not override
    if (_toolbarForcedMode) return;
    var showToolbar = window.scrollY > 60;
    navbar.classList.toggle('navbar-hidden', showToolbar);
    toolbar.classList.toggle('visible', showToolbar);
}

function _activateToolbarForced() {
    // Force sticky toolbar to show (used when editor gains focus)
    _toolbarForcedMode = true;
    var navbar = document.getElementById('mainNavbar');
    var toolbar = document.getElementById('stickyToolbar');
    console.log('[DEBUG] _activateToolbarForced called, navbar:', !!navbar, 'toolbar:', !!toolbar);
    navbar.classList.add('navbar-hidden');
    toolbar.classList.add('visible');
    console.log('[DEBUG] toolbar classes after:', toolbar.className);
}

function _onEditorBlur() {
    // After blur, clear forced state and restore scroll-based visibility
    _toolbarForcedMode = false;
    _updateHeaderVisibility();
}

function _getMostVisibleSection() {
    var vpCenter = window.innerHeight / 2;
    var bestId = null;
    var bestDist = Infinity;
    SECTIONS.forEach(function(sec) {
        var el = document.getElementById(sec.id);
        if (!el) return;
        var card = el.closest('.section-card');
        if (!card) return;
        var rect = card.getBoundingClientRect();
        if (rect.bottom < 0 || rect.top > window.innerHeight) return;
        var dist = Math.abs((rect.top + rect.bottom) / 2 - vpCenter);
        if (dist < bestDist) {
            bestDist = dist;
            bestId = sec.id;
        }
    });
    return bestId;
}

function _onScroll() {
    // When user scrolls past navbar after clicking, hand off to scroll-based visibility
    if (window.scrollY > 60 && _toolbarForcedMode) {
        _toolbarForcedMode = false;
    } else if (window.scrollY <= 60) {
        // Back at top - restore click-mode possible (toolbar hidden until focus)
        _toolbarForcedMode = false;
    }
    _updateHeaderVisibility();

    // Only auto-switch editors when scrolled past navbar
    clearTimeout(_scrollSwitchTimer);
    if (window.scrollY > 60) {
        _scrollSwitchTimer = setTimeout(function() {
            var visibleId = _getMostVisibleSection();
            if (visibleId && visibleId !== _activeSecId) {
                var editor = tinymce.get(visibleId);
                if (editor) {
                    var sx = window.scrollX, sy = window.scrollY;
                    editor.focus();
                    window.scrollTo(sx, sy);
                }
            }
        }, 120);
    }
}

// ===== Init =====
document.addEventListener('DOMContentLoaded', function() {
    // Set navbar spacer height
    var navbar = document.getElementById('mainNavbar');
    var spacer = document.getElementById('navbarSpacer');
    if (navbar && spacer) {
        spacer.style.height = navbar.offsetHeight + 'px';
    }
    // Initialize toolbar visibility on page load
    _updateHeaderVisibility();

    renderSections();
    initEditors();

    // Scroll listener
    window.addEventListener('scroll', _onScroll, { passive: true });

    // Sync the two doc name inputs + update sidebar
    var docName1 = document.getElementById('docName');
    var docName2 = document.getElementById('docNameToolbar');
    if (docName1 && docName2) {
        docName1.addEventListener('input', function() {
            docName2.value = docName1.value;
            _syncDocNameToSidebar();
        });
        docName2.addEventListener('input', function() {
            docName1.value = docName2.value;
            _syncDocNameToSidebar();
        });
    }
});

function _syncDocNameToSidebar() {
    if (!_draftsData || !_draftsData.activeId) return;
    var name = document.getElementById('docName').value.trim();
    _draftsData.docs[_draftsData.activeId].docName = name;
    // Update just the active item's text (avoid full re-render flicker)
    var activeItem = document.querySelector('.sidebar-doc-item.active .sidebar-doc-name');
    if (activeItem) activeItem.textContent = name || '未命名制度';
}

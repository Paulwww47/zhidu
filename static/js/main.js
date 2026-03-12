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
                <textarea id="${sec.id}"></textarea>
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

// ===== Initialize TinyMCE =====
function initEditors() {
    tinymce.init({
        selector: SECTIONS.map(s => '#' + s.id).join(','),
        language: 'zh_CN',
        language_url: '/static/tinymce/langs/zh_CN.js',
        skin_url: '/tinymce/skins/ui/oxide',
        content_css: '/tinymce/skins/content/default/content.min.css',
        height: 450,
        menubar: 'file edit view insert format table',
        plugins: [
            'advlist', 'autolink', 'lists', 'link', 'image', 'charmap', 'preview',
            'searchreplace', 'visualblocks', 'fullscreen',
            'insertdatetime', 'table', 'wordcount'
        ],
        toolbar: 'undo redo | blocks | bold italic underline strikethrough | ' +
                 'forecolor backcolor | alignleft aligncenter alignright alignjustify firstindent | ' +
                 'bullist numlist outdent indent | image table cellvalign | removeformat',
        formats: {
            firstindent: { selector: 'p,div', styles: { 'text-indent': '2em' } }
        },
        setup: function(editor) {
            editor.ui.registry.addToggleButton('firstindent', {
                text: '首行缩进',
                tooltip: '首行缩进 2字符',
                onAction: function() {
                    var node = editor.selection.getNode();
                    var p = node.closest('p,div') || node;
                    if (p && (p.nodeName === 'P' || p.nodeName === 'DIV')) {
                        if (p.style.textIndent) {
                            p.style.textIndent = '';
                        } else {
                            p.style.textIndent = '2em';
                        }
                        editor.nodeChanged();
                    }
                },
                onSetup: function(api) {
                    editor.on('NodeChange', function() {
                        var node = editor.selection.getNode();
                        var p = node.closest('p,div') || node;
                        api.setActive(p && p.style && p.style.textIndent === '2em');
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
                                var selectedCells = editor.dom.select('td.mce-item-selected,th.mce-item-selected');
                                if (selectedCells.length > 0) {
                                    selectedCells.forEach(function(cell) {
                                        cell.style.verticalAlign = 'top';
                                    });
                                } else {
                                    var cell = editor.selection.getNode().closest('td,th');
                                    if (cell) cell.style.verticalAlign = 'top';
                                }
                                editor.nodeChanged();
                            }
                        },
                        {
                            type: 'menuitem',
                            text: '垂直居中',
                            onAction: function() {
                                var selectedCells = editor.dom.select('td.mce-item-selected,th.mce-item-selected');
                                if (selectedCells.length > 0) {
                                    selectedCells.forEach(function(cell) {
                                        cell.style.verticalAlign = 'middle';
                                    });
                                } else {
                                    var cell = editor.selection.getNode().closest('td,th');
                                    if (cell) cell.style.verticalAlign = 'middle';
                                }
                                editor.nodeChanged();
                            }
                        },
                        {
                            type: 'menuitem',
                            text: '底部对齐',
                            onAction: function() {
                                var selectedCells = editor.dom.select('td.mce-item-selected,th.mce-item-selected');
                                if (selectedCells.length > 0) {
                                    selectedCells.forEach(function(cell) {
                                        cell.style.verticalAlign = 'bottom';
                                    });
                                } else {
                                    var cell = editor.selection.getNode().closest('td,th');
                                    if (cell) cell.style.verticalAlign = 'bottom';
                                }
                                editor.nodeChanged();
                            }
                        }
                    ]);
                }
            });
        },
        table_default_styles: {
            'border-collapse': 'collapse',
            'width': '100%'
        },
        table_cell_advtab: true,
        images_upload_url: '/api/upload-image',
        automatic_uploads: true,
        file_picker_types: 'image',
        paste_preprocess: function(plugin, args) {
            var c = args.content;
            // 1. Conditional comments: <!--[if ...]>...<![endif]-->
            c = c.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, '');
            // 2. Stray [if !supportLists] / [endif] without HTML comment wrapper
            c = c.replace(/\[if\s+!support\w+\]/gi, '');
            c = c.replace(/\[endif\]/gi, '');
            // 3. XML declarations <?xml ...?>
            c = c.replace(/<\?xml[\s\S]*?\?>/gi, '');
            // 4. <style> blocks (Word embeds massive mso style blocks)
            c = c.replace(/<style[\s\S]*?<\/style>/gi, '');
            // 5. Namespace tags: <o:p>, </o:p>, <v:shape>, <w:wrap>, <st1:*>, <m:*> etc.
            c = c.replace(/<\/?\w+:[^>]*>/gi, '');
            // 6. xmlns:* attributes
            c = c.replace(/\s*xmlns:\w+="[^"]*"/gi, '');
            // 7. class="Mso..." attributes
            c = c.replace(/\s*class="Mso[^"]*"/gi, '');
            // 8. mso-* CSS properties inside style attributes
            c = c.replace(/mso-[a-z\-]+\s*:[^;"]*;?/gi, '');
            // 9. Clean up empty style="" left after mso removal
            c = c.replace(/\s*style="\s*"/gi, '');
            // 10. <font> tags (keep content)
            c = c.replace(/<\/?font[^>]*>/gi, '');
            // 11. <span style="mso-spacerun:yes"> &nbsp; </span>
            c = c.replace(/<span[^>]*mso-spacerun[^>]*>[\s\u00a0]*<\/span>/gi, '');
            // 12. Empty <span></span> wrappers
            c = c.replace(/<span><\/span>/gi, '');
            // 13. lang / xml:lang attributes
            c = c.replace(/\s*(xml:)?lang="[^"]*"/gi, '');
            args.content = c;
        },
        content_style: 'body { font-family: SimSun, serif; font-size: 12pt; }',
        branding: false,
        promotion: false,
        license_key: 'gpl',
        convert_urls: false,
        relative_urls: false
    });
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
            body: JSON.stringify({ sections })
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

// ===== Init =====
document.addEventListener('DOMContentLoaded', () => {
    renderSections();
    // Small delay to let textareas be in DOM
    setTimeout(initEditors, 100);
});

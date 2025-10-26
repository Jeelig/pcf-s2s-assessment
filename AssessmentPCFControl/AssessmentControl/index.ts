import { IInputs, IOutputs } from "./generated/ManifestTypes";
import {
    AuditQuestion,
    SubQuestion,
    QuestionCategory,
    AuditTemplate,
    AuditRecord,
    PendingChange,
    OfflineData,
    StorageEntity
} from "./types";

export class AssessmentControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private context: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged: () => void;
    
    private auditId: string;
    private catquestions: QuestionCategory[] = [];
    private allcollapse = false;
    private auditdisabled = false;
    private templateType: AuditTemplate = {} as AuditTemplate;
    private globalEntity: StorageEntity = {};
    private globalTemplate: AuditTemplate = {} as AuditTemplate;
    private globalResult: AuditRecord = {} as AuditRecord;
    
    private isOffline = false;
    private pendingChanges: PendingChange[] = [];
    private offlineData: OfflineData | null = null;
    
    private changedQuestions = new Map<string, AuditQuestion>();
    private changedSubquestions = new Map<string, SubQuestion>();

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
        
        this.checkOnlineStatus();
        
        this.auditId = context.parameters.auditId.raw || "";
        
        if (!this.auditId) {
            this.showError("No Assessment ID is given!");
            return;
        }
        
        this.loadOfflineData();
        this.renderUI();
        this.loadAuditFromInput();
        
        window.addEventListener('online', () => {
            this.handleOnline();
        });
        window.addEventListener('offline', () => {
            this.handleOffline();
        });
    }

    private checkOnlineStatus(): void {
        this.isOffline = !navigator.onLine;
    }

    private handleOnline(): void {
        console.log("Connection restored");
        this.isOffline = false;
        this.updateConnectionStatus(true);
        this.notifyOutputChanged();
    }

    private handleOffline(): void {
        console.log("Connection lost - switching to offline mode");
        this.isOffline = true;
        this.updateConnectionStatus(false);
    }

    private updateConnectionStatus(isOnline: boolean): void {
        const statusDiv = this.container.querySelector("#connectionStatus");
        if (statusDiv) {
            if (isOnline) {
                statusDiv.textContent = "✓ Online - Ready to sync";
                (statusDiv as HTMLElement).className = "connection-status online";
                setTimeout(() => {
                    (statusDiv as HTMLElement).style.display = "none";
                }, 3000);
            } else {
                statusDiv.textContent = "⚠ Offline Mode - Changes saved locally";
                (statusDiv as HTMLElement).className = "connection-status offline";
                (statusDiv as HTMLElement).style.display = "block";
            }
        }
    }

    private loadOfflineData(): void {
        try {
            const stored = localStorage.getItem(`audit_${this.auditId}`);
            if (stored) {
                this.offlineData = JSON.parse(stored) as OfflineData;
                console.log("Loaded offline data for audit:", this.auditId);
            }
            
            const pending = localStorage.getItem(`pending_${this.auditId}`);
            if (pending) {
                this.pendingChanges = JSON.parse(pending) as PendingChange[];
                console.log("Loaded pending changes:", this.pendingChanges.length);
            }
        } catch (e) {
            console.error("Error loading offline data:", e);
        }
    }

    private saveOfflineData(): void {
        try {
            const dataToSave: OfflineData = {
                catquestions: this.catquestions,
                globalResult: this.globalResult,
                globalTemplate: this.globalTemplate,
                templateType: this.templateType,
                lastSync: new Date().toISOString()
            };
            
            localStorage.setItem(`audit_${this.auditId}`, JSON.stringify(dataToSave));
            console.log("Saved offline data for audit:", this.auditId);
        } catch (e) {
            console.error("Error saving offline data:", e);
        }
    }

    private renderUI(): void {
        this.container.innerHTML = `
            <div class="connection-status" id="connectionStatus"></div>
            <div class="alertMessage" id="alertMessage"></div>
            <div class="savedMessage" id="savedMessage">Changes saved locally!</div>
            <div class="s2s-container">
                <div style="text-align:right; margin-bottom: 20px;">
                    ${this.pendingChanges.length > 0 ? 
                        `<span class="pending-badge">${this.pendingChanges.length} pending</span>` : ''}
                    <label class="switch">
                        <input type="checkbox" id="collapseToggle">
                        <span class="slider"></span>
                    </label>
                    <span class="label">Collapse all</span>
                </div>
                <div id="categoriesContainer"></div>
            </div>
        `;

        const collapseToggle = this.container.querySelector("#collapseToggle") as HTMLInputElement;
        if (collapseToggle) {
            collapseToggle.addEventListener("change", () => this.collapseAll(collapseToggle.checked));
        }
        
        this.updateConnectionStatus(!this.isOffline);
    }

    private loadAuditFromInput(): void {
        try {
            const auditDataParam = this.context.parameters.auditData?.raw;
            if (auditDataParam && auditDataParam.length > 0) {
                console.log("Loading audit from Canvas App input");
                console.log("Raw audit data:", auditDataParam.substring(0, 200));  // ADDED: Log first 200 chars
                try {  // ADDED: Try-catch for JSON parsing
                    const data = JSON.parse(auditDataParam) as AuditRecord;
                    console.log("Parsed audit data successfully");  // ADDED: Success log
                    this.processAuditData(data);
                    return;
                } catch (parseError) {  // ADDED: Catch JSON parse errors
                    console.error("JSON parse error:", parseError);
                    this.showError("Failed to parse audit data. Please check the JSON format.");
                    return;
                }
            }
            if (this.offlineData) {
                console.log("Loading from offline storage");
                this.catquestions = this.offlineData.catquestions || [];
                this.globalResult = this.offlineData.globalResult || {} as AuditRecord;
                this.globalTemplate = this.offlineData.globalTemplate || {} as AuditTemplate;
                this.templateType = this.offlineData.templateType || {} as AuditTemplate;
                this.renderCategories();
                return;
            }
            console.warn("No audit data available");  // ADDED: Warning log
            this.showError("No audit data available. Please ensure Canvas App passes data.");
        } catch (error) {
            console.error("Error loading audit:", error);
            this.showError("Failed to load audit data: " + String(error));  // CHANGED: Added error details
        }
    }

    private processAuditData(result: AuditRecord): void {
        this.globalResult = result;
        
        if (result.statuscode === 181910001) {
            this.auditdisabled = true;
        }

        this.globalTemplate = this.templateType;
        this.catquestions = this.orderResult(result);
        this.setResponses();
        this.saveOfflineData();
        this.renderCategories();
    }

    private orderResult(result: AuditRecord): QuestionCategory[] {
        const categoriesMap: Record<string, QuestionCategory> = {};
        const questions = result.nov_audit_nov_auditquestion_audit || [];
        const skuMap: Record<string, AuditQuestion> = {};

        questions.forEach((q: AuditQuestion) => {
            const cat = q.nov_questioncategory;
            if (!cat) return;
            
            const catId = cat.nov_questioncategoryid;

            if (!categoriesMap[catId]) {
                categoriesMap[catId] = {
                    id: catId,
                    name: cat.nov_questioncategory,
                    parentCategoryId: cat._nov_parentquestioncategory_value || undefined,
                    questions: [],
                    subCategories: [],
                    px: 0,
                    percent: 0,
                    nb_answers: 0,
                    hide: false,
                    nov_questioncategoryid: catId,
                    nov_questioncategory: cat.nov_questioncategory
                };
            }

            if (q.nov_questiontype === 181910000) q.sku = true;
            else if (q.nov_questiontype === 181910001) q.nonSku = true;
            else if (q.nov_questiontype === 285050000) q.nonSkuSubQuestion = true;
            else if (q.nov_questiontype === 285050001) q.q1q2 = true;

            if (q.sku && q._nov_question_value) {
                const key = q._nov_question_value;
                if (!skuMap[key]) {
                    skuMap[key] = {
                        ...q,
                        id: key,
                        label: q.nov_auditquestion,
                        score: q.nov_px_score,
                        answer: q.answer,
                        nov_answertype: q.nov_answertype,
                        subQuestions: [],
                        nov_auditquestionid: q.nov_auditquestionid,
                        nov_auditquestion: q.nov_auditquestion,
                        nov_questiontype: q.nov_questiontype
                    };
                }
                skuMap[key].subQuestions = skuMap[key].subQuestions || [];
                skuMap[key].subQuestions!.push({ ...q });
            } else {
                categoriesMap[catId].questions.push({
                    ...q,
                    id: q.nov_auditquestionid,
                    label: q.nov_auditquestion,
                    score: q.nov_px_score,
                    answered: false
                });
            }
        });

        Object.values(skuMap).forEach((group: AuditQuestion) => {
            const catId = group.nov_questioncategory?.nov_questioncategoryid;
            if (catId && categoriesMap[catId]) {
                categoriesMap[catId].questions.push(group);
            }
        });

        return Object.values(categoriesMap);
    }

    private setResponses(): void {
        for (const cat of this.catquestions) {
            for (const q of cat.questions) {
                if (q.nov_questiontype === 181910001) {
                    this.setQuestionResponse(q);
                } else if (q.nov_questiontype === 285050000 || q.nov_questiontype === 285050001) {
                    if (q.subquestions) {
                        q.subquestions.forEach(sub => {
                            this.setSubQuestionResponse(sub);
                        });
                    }
                } else if (q.nov_questiontype === 181910000 && q.subQuestions?.length) {
                    q.subQuestions.forEach((sku: AuditQuestion) => this.setQuestionResponse(sku));
                }
            }
            this.sumCategoryScoring(cat.id);
        }
        this.updateAllPercentages();
    }

    private setQuestionResponse(q: AuditQuestion): void {
        switch (q.nov_answertype) {
            case 181910000:
                if (q.cgi_answer && q.cgi_answer.length > 0) {
                    q.answer = q.cgi_answer;
                    q.answered = true;
                }
                break;
            case 181910001:
                if (q.cgi_answer && q.cgi_answer.length > 0) {
                    q.answer = "Yes";
                    const parsed = parseFloat(q.cgi_answer);
                    if (!isNaN(parsed) && isFinite(parsed)) {
                        q.value = parseInt(q.cgi_answer);
                        q.answered = true;
                    }
                }
                break;
            case 285050000:
                if (q.cgi_answer && q.cgi_answer.length > 0) {
                    q.answer = "Yes";
                    q.value = q.cgi_answer;
                    q.answered = true;
                } else {
                    q.value = "none";
                }
                break;
        }
    }

    private setSubQuestionResponse(sub: SubQuestion): void {
        switch (sub.cgi_answertype) {
            case 181910000:
                if (sub.cgi_answertext && sub.cgi_answertext.length > 0) {
                    sub.answer = sub.cgi_answertext;
                    sub.answered = true;
                }
                break;
            case 181910001:
                if (sub.cgi_numericalanswer !== undefined && sub.cgi_numericalanswer !== null) {
                    sub.value = sub.cgi_numericalanswer;
                    sub.answered = true;
                }
                break;
        }
    }

    private renderCategories(): void {
        const container = this.container.querySelector("#categoriesContainer");
        if (!container) return;

        container.innerHTML = this.catquestions.map(cat => this.renderCategory(cat)).join("");
        this.attachEventListeners();
    }

    private renderCategory(cat: QuestionCategory): string {
        const progressWidth = cat.percent || 0;
        const statusClass = cat.nb_answers >= cat.questions.length ? 'full' : 'partial';
        
        return `
            <div class="s2s-category" data-cat-id="${cat.id}">
                <table style="width:100%">
                    <tbody>
                        <tr>
                            <td style="width:50px">
                                <div class="status_bubble ${statusClass}">${cat.nb_answers}/${cat.questions.length}</div>
                            </td>
                            <td style="font-weight:bold">${cat.name}</td>
                            <td style="width:90px">
                                ${!this.templateType.cgi_ultraselective ? `<div class="blue_bubble">PX: ${cat.px.toFixed(2)}%</div>` : ''}
                            </td>
                            <td style="width:40px">
                                <div class="icon active collapse-btn" data-cat-id="${cat.id}">
                                    ${cat.hide ? 'expand_more' : 'expand_less'}
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" style="padding: 0px 0 7px 0;">
                                <div class="gray_line"></div>
                                ${!cat.hide && !this.allcollapse ? `
                                    <div class="progress-container">
                                        <div class="progress-bar" style="width: ${progressWidth}px"></div>
                                    </div>
                                ` : ''}
                            </td>
                        </tr>
                        ${!cat.hide && !this.allcollapse ? `
                            <tr>
                                <td colspan="4">
                                    <ul class="question_list">
                                        ${cat.questions.map((q: AuditQuestion) => this.renderQuestion(q, cat)).join('')}
                                    </ul>
                                </td>
                            </tr>
                        ` : ''}
                    </tbody>
                </table>
            </div>
        `;
    }

    private renderQuestion(q: AuditQuestion, cat: QuestionCategory): string {
        const answeredClass = q.sku || q.nonSkuSubQuestion || q.q1q2 ? 'answered-bold' : '';
        
        return `
            <li>
                <table class="tab_question_line">
                    <tbody>
                        <tr>
                            <td>
                                <span class="material-icons quest info-icon" data-q-id="${q.id}">info_outline</span>
                                <span class="q_question ${answeredClass}">${q.nov_auditquestion || q.label}</span>
                                ${q.show_desc ? `<div class="q_description">${q.nov_tdq_description || ''}</div>` : ''}
                            </td>
                            ${this.renderAnswerInput(q, cat)}
                            ${this.renderSubquestionBubble(q)}
                            ${this.renderQ1Q2Score(q)}
                        </tr>
                        ${this.renderSubquestions(q)}
                    </tbody>
                </table>
            </li>
        `;
    }

    private renderAnswerInput(q: AuditQuestion, cat: QuestionCategory): string {
        if (q.nov_answertype === 181910001 && q.nov_questiontype === 181910001) {
            return `
                <td style="width:60px">
                    <input type="number" placeholder="0.00" class="inputnb question-input" 
                           data-q-id="${q.id}" data-cat-id="${cat.id}" data-type="number"
                           value="${q.value || ''}" ${this.auditdisabled ? 'disabled' : ''} min="0">
                </td>
            `;
        } else if (q.nov_answertype === 285050000 && q.nov_questiontype === 181910001) {
            return `
                <td style="width:120px; padding-right:10px;">
                    <select class="inputslct question-input" data-q-id="${q.id}" 
                            data-cat-id="${cat.id}" data-type="listoption"
                            ${this.auditdisabled ? 'disabled' : ''}>
                        <option value="none">- Choose</option>
                        ${(q.list_options || []).map((opt) => 
                            `<option value="${opt.cgi_name}" ${q.value === opt.cgi_name ? 'selected' : ''}>${opt.cgi_name}</option>`
                        ).join('')}
                    </select>
                </td>
            `;
        } else if (q.nov_answertype === 181910000 && q.nov_questiontype === 181910001) {
            return `
                <td style="width:110px;">
                    <div class="custom-radio">
                        <input type="radio" name="yes_no_${q.id}" value="Yes" 
                               class="question-radio" data-q-id="${q.id}" data-cat-id="${cat.id}"
                               ${q.answer === 'Yes' ? 'checked' : ''} 
                               ${this.auditdisabled ? 'disabled' : ''}>
                        <label>Yes</label>
                        <input type="radio" name="yes_no_${q.id}" value="No"
                               class="question-radio" data-q-id="${q.id}" data-cat-id="${cat.id}"
                               ${q.answer === 'No' ? 'checked' : ''}
                               ${this.auditdisabled ? 'disabled' : ''}>
                        <label>No</label>
                    </div>
                </td>
            `;
        }
        return '<td></td>';
    }

    private renderSubquestionBubble(q: AuditQuestion): string {
        if (q.nonSkuSubQuestion && q.subquestions && q.subquestions.length > 0) {
            const answeredCount = q.subquestions.filter(sub => sub.answered).length;
            const totalCount = q.subquestions.length;
            const bubbleClass = answeredCount === totalCount ? 'green_cl' : 'red_cl';
            
            return `
                <td style="width:100px; vertical-align: top;">
                    <div class="subquestion_bubble ${bubbleClass}" data-q-id="${q.id}" data-toggle="sub">
                        ${answeredCount} / ${totalCount}
                    </div>
                </td>
            `;
        }
        
        if (q.q1q2 && q.subquestions && q.subquestions.length > 0) {
            const answeredCount = q.subquestions.filter(sub => sub.answered).length;
            const totalCount = q.subquestions.length;
            const bubbleClass = answeredCount === totalCount ? 'green_cl' : 'red_cl';
            
            return `
                <td style="width:100px; vertical-align: top;">
                    <div class="subquestion_bubble ${bubbleClass}" data-q-id="${q.id}" data-toggle="sub">
                        ${answeredCount} / ${totalCount}
                    </div>
                </td>
            `;
        }
        
        if (q.sku && q.subQuestions && q.subQuestions.length > 0) {
            const answeredCount = q.subQuestions.filter(sub => sub.answered).length;
            const totalCount = q.subQuestions.length;
            const bubbleClass = answeredCount === totalCount ? 'green_cl' : 'red_cl';
            
            return `
                <td style="width:100px; vertical-align: top;">
                    <div class="subquestion_bubble ${bubbleClass}" data-q-id="${q.id}" data-toggle="sub">
                        ${answeredCount} / ${totalCount}
                    </div>
                </td>
            `;
        }
        
        return '';
    }

    private renderQ1Q2Score(q: AuditQuestion): string {
        if (q.q1q2) {
            return `
                <td style="width:100px; vertical-align:top; text-align:right; padding-top:3px;">
                    <div class="nb_block">${q.nov_scored || '&nbsp;'}</div>
                    <div class="nb_block" style="margin-left:-4px;">%</div>
                </td>
            `;
        }
        return '';
    }

    private renderSubquestions(q: AuditQuestion): string {
        const showSub = q.showsub || false;
        const colspan = q.q1q2 ? '3' : '2';
        
        if (!showSub) {
            return '';
        }

        if ((q.nonSkuSubQuestion || q.q1q2) && q.subquestions && q.subquestions.length > 0) {
            const subquestionsHtml = q.subquestions
                .sort((a, b) => (a.cgi_questionflow || 0) - (b.cgi_questionflow || 0))
                .map(sub => this.renderSubquestionRow(sub, q))
                .join('');
            
            return `
                <tr>
                    <td colspan="${colspan}" style="padding:20px 0">
                        ${subquestionsHtml}
                    </td>
                </tr>
            `;
        }

        if (q.sku && q.subQuestions && q.subQuestions.length > 0) {
            const skuHtml = q.subQuestions
                .sort((a, b) => (a.nov_questionflow || 0) - (b.nov_questionflow || 0))
                .map(sku => this.renderSkuRow(sku, q))
                .join('');
            
            return `
                <tr>
                    <td colspan="${colspan}" style="padding:20px 0">
                        ${skuHtml}
                    </td>
                </tr>
            `;
        }

        return '';
    }

    private renderSubquestionRow(sub: SubQuestion, parentQuestion: AuditQuestion): string {
        return `
            <div class="sub_question">
                <table>
                    <tbody>
                        <tr>
                            <td>${sub.cgi_name}</td>
                            ${this.renderSubquestionInput(sub, parentQuestion)}
                        </tr>
                    </tbody>
                </table>
            </div>
        `;
    }

    private renderSubquestionInput(sub: SubQuestion, parentQuestion: AuditQuestion): string {
        if (sub.cgi_answertype === 181910000) {
            return `
                <td style="width:110px;">
                    <div class="custom-radio">
                        <input type="radio" name="yes_no_${sub.cgi_auditsubquestionid}" 
                               value="Yes" class="subquestion-radio" 
                               data-sub-id="${sub.cgi_auditsubquestionid}" 
                               data-q-id="${parentQuestion.id}"
                               ${sub.answer === 'Yes' ? 'checked' : ''} 
                               ${this.auditdisabled ? 'disabled' : ''}>
                        <label>Yes</label>
                        <input type="radio" name="yes_no_${sub.cgi_auditsubquestionid}" 
                               value="No" class="subquestion-radio" 
                               data-sub-id="${sub.cgi_auditsubquestionid}" 
                               data-q-id="${parentQuestion.id}"
                               ${sub.answer === 'No' ? 'checked' : ''} 
                               ${this.auditdisabled ? 'disabled' : ''} 
                               style="margin-left:5px;">
                        <label>No</label>
                    </div>
                </td>
            `;
        } else if (sub.cgi_answertype === 181910001) {
            return `
                <td style="width:100px;">
                    <input type="number" placeholder="0.00" class="inputnb subquestion-input" 
                           data-sub-id="${sub.cgi_auditsubquestionid}" 
                           data-q-id="${parentQuestion.id}"
                           value="${sub.value || ''}" 
                           ${this.auditdisabled ? 'disabled' : ''} 
                           min="0">
                </td>
            `;
        }
        return '<td></td>';
    }

    private renderSkuRow(sku: AuditQuestion, parentQuestion: AuditQuestion): string {
        const details = parentQuestion.sku 
            ? ` <span style="color:#585858ad">/ Target: ${sku.nov_target_formatted || ''} / Product Range: ${sku._nov_productrange_value_formatted || ''} / Life Stage: ${sku.nov_lifestage_formatted || ''} / Stock Weight: ${sku.stockweight_formatted || ''}</span>`
            : '';

        return `
            <div class="sub_question">
                <table>
                    <tbody>
                        <tr>
                            <td>
                                ${sku.nov_auditquestion}${details}
                            </td>
                            ${this.renderSkuInput(sku, parentQuestion)}
                        </tr>
                    </tbody>
                </table>
            </div>
        `;
    }

    private renderSkuInput(sku: AuditQuestion, parentQuestion: AuditQuestion): string {
        if (sku.nov_answertype === 181910000) {
            return `
                <td style="width:110px;">
                    <div class="custom-radio">
                        <input type="radio" name="yes_no_${sku.nov_auditquestionid}" 
                               value="Yes" class="sku-radio" 
                               data-sku-id="${sku.nov_auditquestionid}" 
                               data-parent-id="${parentQuestion.id}"
                               ${sku.answer === 'Yes' ? 'checked' : ''} 
                               ${this.auditdisabled ? 'disabled' : ''}>
                        <label>Yes</label>
                        <input type="radio" name="yes_no_${sku.nov_auditquestionid}" 
                               value="No" class="sku-radio" 
                               data-sku-id="${sku.nov_auditquestionid}" 
                               data-parent-id="${parentQuestion.id}"
                               ${sku.answer === 'No' ? 'checked' : ''} 
                               ${this.auditdisabled ? 'disabled' : ''} 
                               style="margin-left:5px;">
                        <label>No</label>
                    </div>
                </td>
            `;
        } else if (sku.nov_answertype === 181910001) {
            return `
                <td style="width:100px;">
                    <span>${sku.isstockitem || ''}</span>
                    <input type="number" placeholder="0.00" class="inputnb sku-input" 
                           data-sku-id="${sku.nov_auditquestionid}" 
                           data-parent-id="${parentQuestion.id}"
                           value="${sku.value || ''}" 
                           ${this.auditdisabled ? 'disabled' : ''} 
                           min="0">
                </td>
            `;
        }
        return '<td></td>';
    }

    private attachEventListeners(): void {
        this.container.querySelectorAll(".collapse-btn").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const catId = (e.target as HTMLElement).getAttribute("data-cat-id");
                if (catId) this.toggleCategory(catId);
            });
        });

        this.container.querySelectorAll(".question-input").forEach(input => {
            input.addEventListener("change", (e) => this.handleAnswerChange(e));
        });

        this.container.querySelectorAll(".question-radio").forEach(radio => {
            radio.addEventListener("change", (e) => this.handleAnswerChange(e));
        });

        this.container.querySelectorAll(".info-icon").forEach(icon => {
            icon.addEventListener("click", (e) => {
                const qId = (e.target as HTMLElement).getAttribute("data-q-id");
                if (qId) this.toggleDescription(qId);
            });
        });

        this.container.querySelectorAll("[data-toggle='sub']").forEach(bubble => {
            bubble.addEventListener("click", (e) => {
                const qId = (e.target as HTMLElement).getAttribute("data-q-id");
                if (qId) this.toggleSubquestions(qId);
            });
        });

        this.container.querySelectorAll(".subquestion-input").forEach(input => {
            input.addEventListener("change", (e) => this.handleSubquestionChange(e));
        });

        this.container.querySelectorAll(".subquestion-radio").forEach(radio => {
            radio.addEventListener("change", (e) => this.handleSubquestionChange(e));
        });

        this.container.querySelectorAll(".sku-input").forEach(input => {
            input.addEventListener("change", (e) => this.handleSkuChange(e));
        });

        this.container.querySelectorAll(".sku-radio").forEach(radio => {
            radio.addEventListener("change", (e) => this.handleSkuChange(e));
        });
    }

    private toggleCategory(catId: string): void {
        const cat = this.catquestions.find(c => c.id === catId);
        if (cat) {
            cat.hide = !cat.hide;
            this.saveOfflineData();
            this.renderCategories();
        }
    }

    private toggleDescription(qId: string): void {
        for (const cat of this.catquestions) {
            const q = cat.questions.find((question: AuditQuestion) => question.id === qId);
            if (q) {
                q.show_desc = !q.show_desc;
                this.renderCategories();
                break;
            }
        }
    }

    private toggleSubquestions(qId: string): void {
        for (const cat of this.catquestions) {
            const q = cat.questions.find((question: AuditQuestion) => question.id === qId);
            if (q) {
                q.showsub = !q.showsub;
                this.renderCategories();
                break;
            }
        }
    }

    private handleAnswerChange(e: Event): void {
        const target = e.target as HTMLInputElement;
        const qId = target.getAttribute("data-q-id");
        const catId = target.getAttribute("data-cat-id");
        const type = target.getAttribute("data-type");

        if (!qId || !catId) return;

        const cat = this.catquestions.find(c => c.id === catId);
        if (!cat) return;

        const question = cat.questions.find((q: AuditQuestion) => q.id === qId);
        if (!question) return;

        if (type === "number") {
            question.value = parseFloat(target.value) || 0;
            question.answer = String(question.value);
            question.answered = question.value > 0;
        } else if (type === "listoption") {
            question.value = target.value;
            question.answer = "Yes";
            question.answered = target.value !== "none";
        } else {
            question.answer = target.value;
            question.answered = true;
        }

        this.changedQuestions.set(question.nov_auditquestionid, question);
        
        this.updatePercentage(question, cat);
        this.calculateScores(question, cat);
        this.saveOfflineData();
        this.renderCategories();
        this.showSuccess();
        this.notifyOutputChanged();
    }

    private handleSubquestionChange(e: Event): void {
        const target = e.target as HTMLInputElement;
        const subId = target.getAttribute("data-sub-id");
        const qId = target.getAttribute("data-q-id");

        if (!subId || !qId) return;

        for (const cat of this.catquestions) {
            const question = cat.questions.find((q: AuditQuestion) => q.id === qId);
            if (question && question.subquestions) {
                const subquestion = question.subquestions.find(sub => sub.cgi_auditsubquestionid === subId);
                if (subquestion) {
                    if (target.type === 'number') {
                        subquestion.value = parseFloat(target.value) || 0;
                        subquestion.answer = String(subquestion.value);
                        subquestion.answered = subquestion.value > 0;
                    } else {
                        subquestion.answer = target.value;
                        subquestion.answered = true;
                    }

                    this.changedSubquestions.set(subquestion.cgi_auditsubquestionid, subquestion);

                    if (question.q1q2 && question.subquestions.length >= 2) {
                        this.computeQ2Q1Calculation(question);
                        this.changedQuestions.set(question.nov_auditquestionid, question);
                    }

                    const allAnswered = question.subquestions.every(sub => sub.answered);
                    if (allAnswered) {
                        question.answered = true;
                        this.updatePercentage(question, cat);
                    }

                    this.calculateScores(question, cat);
                    this.saveOfflineData();
                    this.renderCategories();
                    this.showSuccess();
                    this.notifyOutputChanged();
                    break;
                }
            }
        }
    }

    private handleSkuChange(e: Event): void {
        const target = e.target as HTMLInputElement;
        const skuId = target.getAttribute("data-sku-id");
        const parentId = target.getAttribute("data-parent-id");

        if (!skuId || !parentId) return;

        for (const cat of this.catquestions) {
            const parentQuestion = cat.questions.find((q: AuditQuestion) => q.id === parentId);
            if (parentQuestion && parentQuestion.subQuestions) {
                const skuQuestion = parentQuestion.subQuestions.find(sku => sku.nov_auditquestionid === skuId);
                if (skuQuestion) {
                    if (target.type === 'number') {
                        skuQuestion.value = parseFloat(target.value) || 0;
                        skuQuestion.answer = String(skuQuestion.value);
                        skuQuestion.answered = skuQuestion.value > 0;
                    } else {
                        skuQuestion.answer = target.value;
                        skuQuestion.answered = true;
                    }

                    this.changedQuestions.set(skuQuestion.nov_auditquestionid, skuQuestion);

                    const allAnswered = parentQuestion.subQuestions.every(sub => sub.answered);
                    if (allAnswered) {
                        parentQuestion.answered = true;
                        this.updatePercentage(parentQuestion, cat);
                    }

                    this.calculateScores(parentQuestion, cat);
                    this.saveOfflineData();
                    this.renderCategories();
                    this.showSuccess();
                    this.notifyOutputChanged();
                    break;
                }
            }
        }
    }

    private computeQ2Q1Calculation(question: AuditQuestion): void {
        if (!question.subquestions || question.subquestions.length < 2) return;

        const sortedSubs = [...question.subquestions].sort((a, b) => 
            (a.cgi_questionflow || 0) - (b.cgi_questionflow || 0)
        );

        const q1Value = sortedSubs[0].cgi_numericalanswer || sortedSubs[0].value || 0;
        const q2Value = sortedSubs[1].cgi_numericalanswer || sortedSubs[1].value || 0;

        if (q1Value !== 0) {
            question.nov_scored = parseFloat(((Number(q2Value) / Number(q1Value)) * 100).toFixed(2));
        } else {
            question.nov_scored = 0;
        }
    }

    private calculateScores(question: AuditQuestion, category: QuestionCategory): void {
        this.calculateQuestionScore(question);
        this.sumCategoryScoring(category.id);
        
        if (question.sku) {
            this.calculateAllSkuScores();
        }
        
        this.updateAllAuditScores();
        this.checkCACompliance();
    }

    private calculateQuestionScore(question: AuditQuestion): void {
        if (!question.scoring_rules || question.scoring_rules.length === 0) {
            question.nov_px_score = 0;
            return;
        }

        switch (question.nov_questiontype) {
            case 181910001:
                if (question.nov_answertype === 181910000) {
                    question.nov_px_score = (question.answer === "Yes")
                        ? question.scoring_rules[0].nov_target
                        : 0;
                } else if (question.nov_answertype === 181910001) {
                    question.nov_px_score = this.numericalCalculation(question);
                } else if (question.nov_answertype === 285050000) {
                    question.nov_px_score = this.listOptionCalculation(question);
                }
                break;
                
            case 285050001:
                question.nov_px_score = this.numericalCalculation(question);
                break;

            case 181910000:
                this.calculateSKUPerfectXScore(question);
                break;
        }
    }

    private numericalCalculation(question: AuditQuestion): number {
        const rules = question.scoring_rules || [];
        const sortedRules = [...rules].sort((a, b) => a.nov_threshold - b.nov_threshold);
        const answerValue = question.nov_scored || (typeof question.answer === 'string' ? parseFloat(question.answer) : (question.answer || 0));
        
        for (let i = 0; i < sortedRules.length; i++) {
            const threshold = sortedRules[i].nov_threshold;
            const target = sortedRules[i].nov_target;
            
            if (i === 0 && answerValue <= threshold) {
                return target;
            }
            
            const nextThreshold = sortedRules[i + 1]?.nov_threshold;
            if (i < sortedRules.length - 1 && answerValue > threshold && answerValue <= nextThreshold) {
                return sortedRules[i + 1].nov_target;
            }
            
            if (i === sortedRules.length - 1 && answerValue > threshold) {
                return 0;
            }
        }
        
        return 0;
    }

    private listOptionCalculation(question: AuditQuestion): number {
        if (!question.list_options || !question.scoring_rules) return 0;
        
        const selectedOption = question.list_options.find(opt => opt.cgi_name === question.cgi_answer);
        if (!selectedOption) return 0;
        
        const matchedRule = question.scoring_rules.find(rule => rule.nov_threshold === selectedOption.cgi_value);
        return matchedRule?.nov_target || 0;
    }

    private calculateSKUPerfectXScore(question: AuditQuestion): void {
        const questionScoringRules = question.scoring_rules;
        
        if (!questionScoringRules || questionScoringRules.length === 0) {
            question.nov_px_score = 0;
            return;
        }

        if (!question.subQuestions || question.subQuestions.length === 0) {
            question.nov_px_score = 0;
            return;
        }

        const yesAnswers = question.subQuestions.filter(item => item.answer === "Yes").length;
        const totalQuestions = question.subQuestions.length;
        const answeredValue = (yesAnswers / totalQuestions) * 100;

        let selectedScoringRule = questionScoringRules[0];

        if (questionScoringRules.length > 1) {
            const sortedRules = [...questionScoringRules].sort((sr1, sr2) => sr1.nov_threshold - sr2.nov_threshold);
            selectedScoringRule = sortedRules.find(scoringRule => answeredValue >= scoringRule.nov_threshold) || sortedRules[0];
        }

        if (selectedScoringRule) {
            const threshold = parseFloat(String(selectedScoringRule.nov_threshold));
            const target = parseFloat(String(selectedScoringRule.nov_target));
            const weighted = selectedScoringRule.nov_weighted === true;

            const thresholdGoal = Math.ceil((threshold / 100) * totalQuestions);

            let score = 0;
            if (yesAnswers >= thresholdGoal) {
                if (weighted && totalQuestions !== 0) {
                    score = (yesAnswers / totalQuestions) * target;
                } else {
                    score = target;
                }
            }
            question.nov_px_score = parseFloat(score.toFixed(2));
        } else {
            question.nov_px_score = 0;
        }
    }

    private calculateAllSkuScores(): void {
        this.mustHaveDog();
        this.mustHaveCat();
        this.totalMustHave();
        this.totalDog();
        this.totalCat();
        this.nextBestDog();
        this.nextBestCat();
        this.totalNextBest();
        this.dogDry();
        this.dogWet();
        this.catDry();
        this.catWet();
        this.catReportingRange();
        this.dogReportingRange();
        this.reportingRange();
        this.catTerritory();
        this.dogTerritory();
        this.territory();
        this.otherCat();
        this.otherDog();
    }

    private mustHaveDog(): void {
        let value = 0;
        let totalDog = 0;
        let answeredDog = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterDog = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050000 && item.nov_target_formatted === "Dog"
                    );
                    const filterDogYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050000 && item.nov_target_formatted === "Dog" && item.answer === "Yes"
                    );
                    
                    totalDog += filterDog.length;
                    answeredDog += filterDogYes.length;
                }
            });
        });
        
        if (answeredDog !== 0 && totalDog !== 0) {
            value = (answeredDog / totalDog) * 100;
        }
        
        this.globalEntity.cgi_musthavedog = value;
    }

    private mustHaveCat(): void {
        let value = 0;
        let totalCat = 0;
        let answeredCat = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterCat = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050000 && item.nov_target_formatted === "Cat"
                    );
                    const filterCatYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050000 && item.nov_target_formatted === "Cat" && item.answer === "Yes"
                    );
                    
                    totalCat += filterCat.length;
                    answeredCat += filterCatYes.length;
                }
            });
        });
        
        if (answeredCat !== 0 && totalCat !== 0) {
            value = (answeredCat / totalCat) * 100;
        }
        
        this.globalEntity.cgi_musthavecat = value;
    }

    private totalMustHave(): void {
        let totalMustHaveValue = 0;
        let answeredMustHaveValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterMusthave = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050000 && 
                        (item.nov_target_formatted === "Cat" || item.nov_target_formatted === "Dog")
                    );
                    const filterYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050000 && item.answer === "Yes"
                    );
                    totalMustHaveValue += filterMusthave.length;
                    answeredMustHaveValue += filterYes.length;
                }
            });
        });
        
        const value = answeredMustHaveValue !== 0 && totalMustHaveValue !== 0
            ? (answeredMustHaveValue / totalMustHaveValue) * 100
            : 0;
        
        this.globalEntity.cgi_totalmusthave = value;
    }

    private totalDog(): void {
        let totalDogValue = 0;
        let answeredValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterDog = q.subQuestions.filter(item => item.nov_target_formatted === "Dog");
                    const filterAnsweredDog = q.subQuestions.filter(item => 
                        item.nov_target_formatted === "Dog" && item.answer === "Yes"
                    );
                    totalDogValue += filterDog.length;
                    answeredValue += filterAnsweredDog.length;
                }
            });
        });

        const value = answeredValue !== 0 && totalDogValue !== 0
            ? (answeredValue / totalDogValue) * 100
            : 0;
        
        this.globalEntity.cgi_totaldog = value;
    }

    private totalCat(): void {
        let totalCatValue = 0;
        let answeredValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterCat = q.subQuestions.filter(item => item.nov_target_formatted === "Cat");
                    const filterAnsweredCat = q.subQuestions.filter(item => 
                        item.nov_target_formatted === "Cat" && item.answer === "Yes"
                    );
                    totalCatValue += filterCat.length;
                    answeredValue += filterAnsweredCat.length;
                }
            });
        });

        const value = answeredValue !== 0 && totalCatValue !== 0
            ? (answeredValue / totalCatValue) * 100
            : 0;
        
        this.globalEntity.cgi_totalcat = value;
    }

    private nextBestDog(): void {
        let totalValue = 0;
        let answeredValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterNBestDog = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050001 && item.nov_target_formatted === "Dog"
                    );
                    const filterNBestDogYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050001 && item.nov_target_formatted === "Dog" && item.answer === "Yes"
                    );
                    
                    totalValue += filterNBestDog.length;
                    answeredValue += filterNBestDogYes.length;
                }
            });
        });
        
        const value = answeredValue !== 0 && totalValue !== 0
            ? (answeredValue / totalValue) * 100
            : 0;
            
        this.globalEntity.cgi_nextbestdog = value;
    }

    private nextBestCat(): void {
        let totalValue = 0;
        let answeredValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterNBestCat = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050001 && item.nov_target_formatted === "Cat"
                    );
                    const filterNBestCatYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050001 && item.nov_target_formatted === "Cat" && item.answer === "Yes"
                    );
                    
                    totalValue += filterNBestCat.length;
                    answeredValue += filterNBestCatYes.length;
                }
            });
        });
        
        const value = answeredValue !== 0 && totalValue !== 0
            ? (answeredValue / totalValue) * 100
            : 0;
            
        this.globalEntity.cgi_nextbestcat = value;
    }

    private totalNextBest(): void {
        let totalNextBestValue = 0;
        let answeredNextBestValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterNBestCat = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050001 && 
                        (item.nov_target_formatted === "Cat" || item.nov_target_formatted === "Dog")
                    );
                    const filterNBestCatYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050001 && item.answer === "Yes"
                    );
                    totalNextBestValue += filterNBestCat.length;
                    answeredNextBestValue += filterNBestCatYes.length;
                }
            });
        });
        
        const value = answeredNextBestValue !== 0 && totalNextBestValue !== 0
            ? (answeredNextBestValue / totalNextBestValue) * 100
            : 0;
        
        this.globalEntity.cgi_totalnextbest = value;
    }

    private otherCat(): void {
        let totalValue = 0;
        let answeredValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterOtherCat = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050002 && item.nov_target_formatted === "Cat"
                    );
                    const filterOtherCatYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050002 && item.nov_target_formatted === "Cat" && item.answer === "Yes"
                    );
                    
                    totalValue += filterOtherCat.length;
                    answeredValue += filterOtherCatYes.length;
                }
            });
        });
        
        const value = answeredValue !== 0 && totalValue !== 0
            ? (answeredValue / totalValue) * 100
            : 0;
            
        this.globalEntity.cgi_othercat = value;
    }

    private otherDog(): void {
        let totalValue = 0;
        let answeredValue = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterOtherDog = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050002 && item.nov_target_formatted === "Dog"
                    );
                    const filterOtherDogYes = q.subQuestions.filter(item => 
                        item.rc_ranges === 285050002 && item.nov_target_formatted === "Dog" && item.answer === "Yes"
                    );
                    
                    totalValue += filterOtherDog.length;
                    answeredValue += filterOtherDogYes.length;
                }
            });
        });
        
        const value = answeredValue !== 0 && totalValue !== 0
            ? (answeredValue / totalValue) * 100
            : 0;
            
        this.globalEntity.cgi_otherdog = value;
    }

    private dogDry(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterDogDryYes = q.subQuestions.filter(item => 
                        item.nov_type === 181910000 && item.nov_target_formatted === "Dog" && item.answer === "Yes"
                    );
                    value += filterDogDryYes.length;
                }
            });
        });
        
        this.globalEntity.cgi_dogdry = value;
    }

    private dogWet(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterDogWetYes = q.subQuestions.filter(item => 
                        item.nov_type === 181910001 && item.nov_target_formatted === "Dog" && item.answer === "Yes"
                    );
                    value += filterDogWetYes.length;
                }
            });
        });
        
        this.globalEntity.cgi_dogwet = value;
    }

    private catDry(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterCatDryYes = q.subQuestions.filter(item => 
                        item.nov_type === 181910000 && item.nov_target_formatted === "Cat" && item.answer === "Yes"
                    );
                    value += filterCatDryYes.length;
                }
            });
        });
        
        this.globalEntity.cgi_catdry = value;
    }

    private catWet(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterCatWetYes = q.subQuestions.filter(item => 
                        item.nov_type === 181910001 && item.nov_target_formatted === "Cat" && item.answer === "Yes"
                    );
                    value += filterCatWetYes.length;
                }
            });
        });
        
        this.globalEntity.cgi_catwet = value;
    }

    private catReportingRange(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterCatReportingRange = q.subQuestions.filter(item => 
                        item.nov_reportingrange !== null && 
                        item.nov_reportingrange !== undefined &&
                        item.nov_target_formatted === "Cat" && 
                        item.answer === "Yes"
                    );
                    value += filterCatReportingRange.length;
                }
            });
        });
        
        this.globalEntity.cgi_catreportingrange = value;
    }

    private dogReportingRange(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterDogReportingRange = q.subQuestions.filter(item => 
                        item.nov_reportingrange !== null && 
                        item.nov_reportingrange !== undefined &&
                        item.nov_target_formatted === "Dog" && 
                        item.answer === "Yes"
                    );
                    value += filterDogReportingRange.length;
                }
            });
        });
        
        this.globalEntity.cgi_dogreportingrange = value;
    }

    private reportingRange(): void {
        const value = this.globalTemplate.cgi_reportingrange || 0;
        this.globalEntity.cgi_reportingrange = value;
    }

    private catTerritory(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterCatTerritory = q.subQuestions.filter(item => 
                        item.nov_territory !== null && 
                        item.nov_territory !== undefined &&
                        item.nov_target_formatted === "Cat" && 
                        item.answer === "Yes"
                    );
                    value += filterCatTerritory.length;
                }
            });
        });
        
        this.globalEntity.cgi_catterritory = value;
    }

    private dogTerritory(): void {
        let value = 0;
        
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.sku && q.subQuestions) {
                    const filterDogTerritory = q.subQuestions.filter(item => 
                        item.nov_territory !== null && 
                        item.nov_territory !== undefined &&
                        item.nov_target_formatted === "Dog" && 
                        item.answer === "Yes"
                    );
                    value += filterDogTerritory.length;
                }
            });
        });
        
        this.globalEntity.cgi_dogterritory = value;
    }

    private territory(): void {
        const value = this.globalTemplate.cgi_territory || 0;
        this.globalEntity.cgi_territory = value;
    }

    private updateAllAuditScores(): void {
        let totalPx = 0;
        this.catquestions.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.nov_px_score !== null && q.nov_px_score !== undefined) {
                    totalPx += parseFloat(String(q.nov_px_score));
                }
            });
        });
        
        this.globalEntity.nov_perfectxscore = totalPx;
        
        let nbCatAnswered = 0;
        const nbCat = this.catquestions.length;
        
        for (const cat of this.catquestions) {
            if (cat.nb_answers === cat.questions.length) {
                nbCatAnswered++;
            }
        }
        
        if (nbCat === nbCatAnswered && nbCat > 0) {
            this.globalEntity.rc_all_questions_answered = true;
        }
    }

    private checkCACompliance(): void {
        try {
            let isCACompliant = false;
            let isCAOnlineCompliant = false;
            let caQuestionsExist = false;
            let caOnlineQuestionsExist = false;
            let allCAAnswersValid = true;
            let allCAOnlineAnswersValid = true;

            this.catquestions.forEach(cat => {
                cat.questions.forEach(q => {
                    const subQs = q.subquestions || q.subQuestions || [];
                    const hasSubquestions = Array.isArray(subQs) && subQs.length > 0;
                    const questionType = q.nov_questiontype;
                    const novAnswerType = q.nov_answertype;

                    const includeMainCA = q.cgi_includeinuscompliance === true;
                    const includeSubCA = hasSubquestions && subQs.some((sub: SubQuestion | AuditQuestion) => 
                        (sub as SubQuestion).cgi_includeinuscompliance === true
                    );

                    if (includeMainCA || includeSubCA) {
                        caQuestionsExist = true;

                        if (questionType === 285050001 && hasSubquestions) {
                            if (novAnswerType === 181910001) {
                                const invalidSubs = subQs.filter((sub: SubQuestion | AuditQuestion) => {
                                    const val = (sub as SubQuestion).cgi_answertext;
                                    return (
                                        val === null ||
                                        val === undefined ||
                                        val === "0" ||
                                        String(val) === "0" ||
                                        (typeof val === "string" && val.trim() === "")
                                    );
                                });

                                if (invalidSubs.length > 0) {
                                    allCAAnswersValid = false;
                                }
                            } else {
                                const nonCompliantSubs = subQs.filter((sub: SubQuestion | AuditQuestion) =>
                                    (sub.answer || "").toString().trim().toLowerCase() !== "yes"
                                );
                                if (nonCompliantSubs.length > 0) {
                                    allCAAnswersValid = false;
                                }
                            }
                        } else if (!hasSubquestions) {
                            if ((q.answer || "").toString().trim().toLowerCase() !== "yes") {
                                allCAAnswersValid = false;
                            }
                        } else {
                            const relevantSubs = subQs.filter((sub: SubQuestion | AuditQuestion) => 
                                (sub as SubQuestion).cgi_includeinuscompliance === true
                            );
                            const nonCompliantSubs = relevantSubs.filter((sub: SubQuestion | AuditQuestion) =>
                                (sub.answer || "").toString().trim().toLowerCase() !== "yes"
                            );

                            if (nonCompliantSubs.length > 0) {
                                allCAAnswersValid = false;
                            }
                        }
                    }

                    const includeMainOnline = (q as AuditQuestion & { cgi_includeincaonlinecompliance?: boolean }).cgi_includeincaonlinecompliance === true;
                    const includeSubOnline = hasSubquestions && subQs.some((sub: SubQuestion | AuditQuestion) =>
                        (sub as SubQuestion).cgi_includeincaonlinecompliance === true
                    );

                    if (includeMainOnline || includeSubOnline) {
                        caOnlineQuestionsExist = true;

                        if (questionType === 285050001 && hasSubquestions) {
                            if (novAnswerType === 181910001) {
                                const invalidSubs = subQs.filter((sub: SubQuestion | AuditQuestion) => {
                                    const val = (sub as SubQuestion).cgi_answertext;
                                    return (
                                        val === null ||
                                        val === undefined ||
                                        val === "0" ||
                                        String(val) === "0" ||
                                        (typeof val === "string" && val.trim() === "")
                                    );
                                });

                                if (invalidSubs.length > 0) {
                                    allCAOnlineAnswersValid = false;
                                }
                            } else {
                                const nonCompliantSubs = subQs.filter((sub: SubQuestion | AuditQuestion) =>
                                    (sub.answer || "").toString().trim().toLowerCase() !== "yes"
                                );
                                if (nonCompliantSubs.length > 0) {
                                    allCAOnlineAnswersValid = false;
                                }
                            }
                        } else if (!hasSubquestions) {
                            if ((q.answer || "").toString().trim().toLowerCase() !== "yes") {
                                allCAOnlineAnswersValid = false;
                            }
                        } else {
                            const relevantSubs = subQs.filter((sub: SubQuestion | AuditQuestion) =>
                                (sub as SubQuestion).cgi_includeincaonlinecompliance === true
                            );
                            const nonCompliantSubs = relevantSubs.filter((sub: SubQuestion | AuditQuestion) =>
                                (sub.answer || "").toString().trim().toLowerCase() !== "yes"
                            );

                            if (nonCompliantSubs.length > 0) {
                                allCAOnlineAnswersValid = false;
                            }
                        }
                    }
                });
            });

            isCACompliant = caQuestionsExist && allCAAnswersValid;
            isCAOnlineCompliant = isCACompliant && caOnlineQuestionsExist && allCAOnlineAnswersValid;

            this.globalEntity.cgi_ultraselectivecompliant = isCACompliant;
            this.globalEntity.cgi_caonlinecompliant = isCAOnlineCompliant;
        } catch (e) {
            console.error("Error in checkCACompliance:", e);
        }
    }

    private updatePercentage(question: AuditQuestion, category: QuestionCategory): void {
        question.answered = question.value !== 0 && question.value !== undefined;

        if (Array.isArray(question.subquestions)) {
            question.subquestions.forEach((sub: SubQuestion) => {
                sub.answered = sub.value !== 0 && sub.value !== undefined;
            });
        }

        let answeredCount = 0;
        category.questions.forEach((q: AuditQuestion) => {
            const isMainAnswered = q.answered === true;
            let subAnswered = true;

            if (Array.isArray(q.subquestions) && q.subquestions.length > 0) {
                subAnswered = q.subquestions.every((sub: SubQuestion) => sub.answered === true);
            }

            if (isMainAnswered && subAnswered) {
                answeredCount += 1;
            }
        });

        const totalQuestions = category.questions.length;
        const containerWidth = 600;
        const percentage = (answeredCount / totalQuestions) * 100;

        category.percent = Math.round((percentage / 100) * containerWidth);
        category.nb_answers = answeredCount;
    }

    private updateAllPercentages(): void {
        for (const cat of this.catquestions) {
            for (const q of cat.questions) {
                if (q.answered) {
                    this.updatePercentage(q, cat);
                }
            }
        }
    }

    private sumCategoryScoring(catId: string): number {
        const category = this.catquestions.find(cat => cat.id === catId);
        if (!category) return 0;
        
        const sum = category.questions.reduce((total: number, q: AuditQuestion) => {
            return total + (q.nov_px_score || 0);
        }, 0);
        
        category.px = sum;
        return sum;
    }

    private collapseAll(collapse: boolean): void {
        this.allcollapse = collapse;
        this.catquestions.forEach(cat => {
            cat.hide = collapse;
        });
        this.saveOfflineData();
        this.renderCategories();
    }

    private showError(message: string): void {
        const alertDiv = this.container.querySelector("#alertMessage");
        if (alertDiv) {
            alertDiv.textContent = message;
            (alertDiv as HTMLElement).style.display = "block";
            setTimeout(() => {
                (alertDiv as HTMLElement).style.display = "none";
            }, 3000);
        }
    }

    private showSuccess(): void {
        const successDiv = this.container.querySelector("#savedMessage");
        if (successDiv) {
            (successDiv as HTMLElement).style.display = "block";
            setTimeout(() => {
                (successDiv as HTMLElement).style.display = "none";
            }, 2500);
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        this.checkOnlineStatus();
        
        const newAuditId = context.parameters.auditId.raw || "";
        const newAuditData = context.parameters.auditData?.raw;
        
        if (newAuditId && newAuditId !== this.auditId) {
            this.auditId = newAuditId;
            this.loadOfflineData();
            this.loadAuditFromInput();
        } else if (newAuditData && newAuditData.length > 0) {
            this.loadAuditFromInput();
        }
    }

    public getOutputs(): IOutputs {
        return {
            completionPercentage: this.calculateOverallCompletion(),
            totalScore: this.calculateTotalScore(),
            pendingChanges: this.pendingChanges.length,
            changedQuestionsJSON: JSON.stringify(Array.from(this.changedQuestions.values())),
            changedSubquestionsJSON: JSON.stringify(Array.from(this.changedSubquestions.values())),
            auditScoresJSON: JSON.stringify(this.globalEntity)
        };
    }

    private calculateOverallCompletion(): number {
        let totalQuestions = 0;
        let answeredQuestions = 0;

        this.catquestions.forEach(cat => {
            totalQuestions += cat.questions.length;
            answeredQuestions += cat.nb_answers;
        });

        return totalQuestions > 0 ? Math.round((answeredQuestions / totalQuestions) * 100) : 0;
    }

    private calculateTotalScore(): number {
        return this.catquestions.reduce((total, cat) => total + (cat.px || 0), 0);
    }

    public destroy(): void {
        this.saveOfflineData();
    }
}
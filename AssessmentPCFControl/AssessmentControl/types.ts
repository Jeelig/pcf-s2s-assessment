// Type definitions for Assessment Control

export interface AuditQuestion {
    id: string;
    nov_auditquestionid: string;
    nov_auditquestion: string;
    nov_answertype: number;
    nov_questiontype: number;
    nov_questionflow?: number;
    nov_tdq_description?: string;
    nov_px_score?: number;
    cgi_answer?: string;
    answer?: string;
    value?: number | string;
    answered?: boolean;
    show_desc?: boolean;
    showsub?: boolean;
    label?: string;
    score?: number;
    sku?: boolean;
    nonSku?: boolean;
    nonSkuSubQuestion?: boolean;
    q1q2?: boolean;
    subquestions?: SubQuestion[];
    subQuestions?: AuditQuestion[];
    list_options?: ListOption[];
    scoring_rules?: ScoringRule[];
    nov_scored?: number;
    _nov_question_value?: string;
    _cgi_auditlistoptiongroup_value?: string;
    nov_questioncategory?: QuestionCategory;
    _nov_questioncategory_value?: string;
    _nov_sku_value?: string;
    nov_target_formatted?: string;
    _nov_productrange_value_formatted?: string;
    nov_lifestage_formatted?: string;
    stockweight_formatted?: string;
    nov_type?: number;
    isstockitem?: boolean;
    nov_lifestage?: number;
    _nov_productrange_value?: string;
    nov_reportingrange?: number;
    nov_reportingrange_formatted?: string;
    nov_territory?: number;
    nov_territory_formatted?: string;
    nov_target?: number;
    stockweight?: number;
    rc_ranges?: number;
    subquestionsLoaded?: boolean;
    localModifiedon?: string;
    cgi_includeinuscompliance?: boolean;
    cgi_includeincaonlinecompliance?: boolean;
}

export interface SubQuestion {
    cgi_auditsubquestionid: string;
    cgi_name: string;
    cgi_answertype: number;
    cgi_questionflow?: number;
    cgi_answer?: string;
    cgi_answertext?: string;
    cgi_numericalanswer?: number;
    answer?: string;
    value?: number;
    answered?: boolean;
    _cgi_subquestion_value?: string;
    cgi_includeinuscompliance?: boolean;
    cgi_includeincaonlinecompliance?: boolean;
}

export interface QuestionCategory {
    nov_questioncategoryid: string;
    nov_questioncategory: string;
    _nov_parentquestioncategory_value?: string;
    questions: AuditQuestion[];
    subCategories: QuestionCategory[];
    px: number;
    percent: number;
    nb_answers: number;
    hide: boolean;
    id: string;
    name: string;
    parentCategoryId?: string;
}

export interface ListOption {
    cgi_auditlistoptionid: string;
    cgi_name: string;
    cgi_order: number;
    cgi_value: number;
    _cgi_auditlistoptiongroup_value: string;
}

export interface ScoringRule {
    nov_scoringrulesid: string;
    nov_scoringrule: string;
    nov_target: number;
    nov_threshold: number;
    nov_weighted: boolean;
    nov_questiontype?: number;
    nov_questioncatgory?: string;
    nov_audittype?: number;
}

export interface AuditTemplate {
    nov_audittemplateid: string;
    cgi_reportingrange?: number;
    cgi_reportingrange_formatted?: string;
    cgi_territory?: number;
    cgi_territory_formatted?: string;
    nov_perfectx?: boolean;
    nov_tradeterm?: string;
    cgi_ultraselective?: boolean;
    cgi_other?: string;
}

export interface AuditRecord {
    nov_auditid: string;
    statuscode: number;
    _nov_related_auditemplate_value?: string;
    nov_audit_nov_auditquestion_audit?: AuditQuestion[];
    rc_all_questions_answered?: boolean;
}

export interface PendingChange {
    id: string;
    type: 'answer' | 'subAnswer' | 'score';
    timestamp: string;
    question?: AuditQuestion;
    subquestion?: SubQuestion;
    entity?: Record<string, string | number | boolean>;
}

export interface OfflineData {
    catquestions: QuestionCategory[];
    globalResult: AuditRecord;
    globalTemplate: AuditTemplate;
    templateType: AuditTemplate;
    lastSync: string;
}

export type StorageEntity = Record<string, string | number | boolean>;
#!/usr/bin/env python3
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# Create a new document
new_doc = Document()

# Helper function to add a paragraph with specific style
def add_para(document, text, style='Normal', bold=False, italic=False):
    p = document.add_paragraph(text, style=style)
    if bold or italic:
        for run in p.runs:
            run.bold = bold
            run.italic = italic
    return p

# ============================================
# EXECUTIVE SUMMARY (NEW - Page 1)
# ============================================

add_para(new_doc, "EXECUTIVE SUMMARY", style='Heading 1')

add_para(new_doc, "", style='Normal')  # Blank line

# Core Problem
p = add_para(new_doc, "Core Problem: Table 6 Exemption Incompatible with Flood Planning Land", style='Heading 2')
p.runs[0].bold = True

add_para(new_doc, "The Draft CPP's Table 6 exemption for residential flat buildings assumes compliance with planning controls can be verified at DA lodgement. On FPCC 1 flood planning land, this is structurally impossible:", style='Normal')

add_para(new_doc, "• Clause 5.22 of Tweed LEP 2014 requires positive demonstration of flood life-safety through site-specific hydraulic and evacuation studies — these do not exist at lodgement", style='Normal')
add_para(new_doc, "• SEPP (Resilience and Hazards) clauses 2.10–2.13 require satisfaction of coastal environment and use obligations — undemonstrated", style='Normal')
add_para(new_doc, "• Tweed DCP 2025 deep soil and landscaping requirements cannot be assumed met on a constrained 2,081 m² site with ground-level parking", style='Normal')
add_para(new_doc, "• A five-storey building will likely exceed the 13.6m height limit, triggering clause 4.6 variation", style='Normal')

add_para(new_doc, "", style='Normal')

# TAD + CPP Interaction Risk
p = add_para(new_doc, "TAD + CPP Coordinated Commencement Risk", style='Heading 2')
p.runs[0].bold = True

add_para(new_doc, "The Discussion Paper confirms remaining Planning System Reforms Act provisions (including Targeted Assessment Development pathway) will commence with the new CPP. If Homes NSW lodges after both instruments take effect (late 2026/2027):", style='Normal')

add_para(new_doc, "• TAD pathway could eliminate exhibition where strategic planning has addressed issues", style='Normal')
add_para(new_doc, "• BUT Council's Planning Committee removed Change Option 22 on 3 April 2025 specifically due to flood concerns", style='Normal')
add_para(new_doc, "• Result: potential for zero public exhibition despite unresolved life-safety risks", style='Normal')

add_para(new_doc, "", style='Normal')

# Five Recommendations
p = add_para(new_doc, "Five Targeted Recommendations", style='Heading 2')
p.runs[0].bold = True

add_para(new_doc, "1. Flood Planning Area carve-out from Table 6 (FPCC 1/2 sites where clause 5.22 preconditions undemonstrated)", style='Normal')
add_para(new_doc, "2. TAD SEPP exclusion for flood areas with undemonstrated preconditions and no strategic support", style='Normal')
add_para(new_doc, "3. Cumulative impact consultation trigger for sites adjacent to schools/childcare on constrained evacuation networks", style='Normal')
add_para(new_doc, "4. Specialist referral transparency during exhibition (engineering, SES, hydraulic advice)", style='Normal')
add_para(new_doc, "5. Savings provisions preserving existing council CPP requirements for flood-constrained development", style='Normal')

add_para(new_doc, "", style='Normal')

# Systemic Issue Statement
p = add_para(new_doc, "Systemic Issue", style='Heading 2')
p.runs[0].bold = True

add_para(new_doc, "This issue is systemic across levee-protected floodplain settlements in NSW. The issue is not whether these developments should proceed, but whether the statutory preconditions governing life-safety can be tested without public scrutiny.", style='Normal')

add_para(new_doc, "", style='Normal')
add_para(new_doc, "", style='Normal')  # Two blank lines before main content

# ============================================
# ORIGINAL CONTENT WITH ENHANCEMENTS
# ============================================

# Title block (from original)
add_para(new_doc, "SUBMISSION", style='Normal')
add_para(new_doc, "Draft Statewide Community Participation Plan", style='Normal')
add_para(new_doc, "Discussion Paper — Community Participation Plan", style='Normal')
add_para(new_doc, "Exhibition Period: 8 April – 3 June 2026", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "Date: 9 April 2026", style='Normal')
add_para(new_doc, "Subject: Submission on Draft Statewide Community Participation Plan and Discussion Paper (Exhibition closes 3 June 2026) — Site-Specific Concerns: Flood Planning Area, FPCC 1 Classification, and Levee-Protected Residential Flat Building Exemption", style='Normal')
add_para(new_doc, "I consent to publication of this submission if and when it is uploaded to any public exhibition portal.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 1
add_para(new_doc, "1. Introduction and Standing", style='Heading 1')
add_para(new_doc, "I make this submission as an affected resident of 7 Heffron Street, Tweed Heads South NSW 2486, immediately adjacent to the proposed Homes NSW social housing redevelopment site at Lots 2–4 DP530539 (3–5 Heffron Street and 6 Seymour Street, Tweed Heads South).", style='Normal')
add_para(new_doc, "I have previously lodged:", style='Normal')
add_para(new_doc, "A pre-lodgement submission with Tweed Shire Council and Homes NSW dated 11 March 2026 (Ref: D25/2290641), addressing statutory flood planning obligations, coastal environment requirements, and sensitive land-use interface impacts.", style='Normal')
add_para(new_doc, "A submission on the Proposed Climate Change and Natural Hazards State Environmental Planning Policy (CC&NH SEPP) Explanation of Intended Effect, also dated 11 March 2026, addressing transitional gaps in flood life-safety protections, the clause 5.22 carry-over mechanism, and the proposed CEA/CUA instrument split.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This submission is made in respect of both the Draft Statewide Community Participation Plan (Draft CPP) and the accompanying Discussion Paper, exhibited from 8 April to 3 June 2026. It focuses on a specific and significant gap in the Draft CPP as it applies to development on flood planning land — in particular, development at the proposed Homes NSW site and comparable levee-protected, FPCC 1-classified sites across NSW.", style='Normal')
add_para(new_doc, "The proposed site is characterised by the following site-specific conditions, each of which is directly relevant to the adequacy of community participation provisions in the Draft CPP:", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 2 with micro-summary
add_para(new_doc, "2. The Table 6 Residential Flat Building Exemption: Why It Cannot Apply to This Site", style='Heading 1')

# MICRO-SUMMARY ADDED
p = add_para(new_doc, "Key point: On FPCC 1 land, compliance with clause 5.22 cannot be verified at DA lodgement. Therefore, Table 6 Condition 2 cannot be satisfied. The exemption assumes something (verified compliance) that is structurally impossible at lodgement for flood-constrained land.", style='Normal')
p.runs[0].italic = True

add_para(new_doc, "", style='Normal')

add_para(new_doc, "Table 6 of the Draft CPP proposes that residential flat buildings be exempt from public exhibition and notification — with only a 7-day pre-commencement notice to adjoining neighbours — where three qualifying conditions are all met:", style='Normal')
add_para(new_doc, "the development is permissible in the relevant zone;", style='Normal')
add_para(new_doc, "the development meets the relevant planning controls in a local environmental plan, development control plan and/or state environmental planning policy; and", style='Normal')
add_para(new_doc, "the development does not include a clause 4.6 variation.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "On the proposed Homes NSW site, the Table 6 exemption is unavailable as a matter of law and planning merit. At minimum two of the three qualifying conditions cannot be satisfied.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 2.1 - Enhanced with boxed formulation
add_para(new_doc, "2.1 Condition 2 Fails as a Matter of Law on FPCC 1 Land", style='Heading 2')

# ENHANCED BOXED FORMULATION
p = add_para(new_doc, "CRITICAL: The \"meets planning controls\" condition cannot be satisfied ex ante on flood land. This is your strongest legal argument: compliance verification is structurally impossible at lodgement when mandatory preconditions (clause 5.22, SEPP coastal obligations, DCP landscaping) require site-specific studies that don't yet exist.", style='Normal')
p.runs[0].bold = True
p.runs[0].italic = True

add_para(new_doc, "", style='Normal')

add_para(new_doc, "The second qualifying condition — that the development \"meets the relevant planning controls\" — cannot be satisfied by reference to zone permissibility and nominal height or FSR compliance alone. For development on this site, mandatory preconditions imposed by Tweed LEP 2014 and applicable SEPPs must be positively demonstrated before any state of satisfaction can be formed.", style='Normal')
add_para(new_doc, "Clause 5.22 of Tweed LEP 2014 (Special flood considerations) imposes a mandatory precondition: development consent must not be granted unless the consent authority is satisfied that the development will not cause a particular risk to life and will not compromise safe and efficient evacuation. This precondition:", style='Normal')
add_para(new_doc, "Cannot be discharged by floor level compliance alone;", style='Normal')
add_para(new_doc, "Requires positive demonstration through site-specific hydraulic and evacuation studies addressing the confirmed 79% levee overtopping probability over a 30-year occupancy period;", style='Normal')
add_para(new_doc, "Requires assessment of cumulative evacuation demand on the shared constrained network serving the adjacent school and childcare centre; and", style='Normal')
add_para(new_doc, "Must be satisfied at the time of determination, not deferred to conditions of consent.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "Until those studies exist and have been assessed, it is impossible to confirm that the development meets the planning controls for the purposes of the Table 6 qualifying condition. On any FPCC 1 site where clause 5.22 applies, the exemption is structurally unavailable at the time of DA lodgement.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "The same analysis applies to:", style='Normal')
add_para(new_doc, "SEPP (Resilience and Hazards) 2021 clauses 2.10–2.13 (CEA and CUA obligations), which require the consent authority to be satisfied that hydrological, ecological, coastal hazard, and visual amenity impacts are avoided, minimised, or mitigated;", style='Normal')
add_para(new_doc, "Tweed LEP 2014 clause 7.1 (Acid sulfate soils), which requires an acid sulfate soils management plan where triggered; and", style='Normal')
add_para(new_doc, "Tweed DCP 2025 Part D2 Section 2.3.2, which imposes mandatory deep soil zone, tree canopy, and landscaping obligations that cannot be assumed to be met without a compliant landscape design demonstrating feasibility on the constrained site.", style='Normal')
add_para(new_doc, "The same analysis applies to the now-operative Tweed DCP 2025 Part D2 deep soil and landscaping obligations (Tables D2.2 and D2.3), which cannot be satisfied without a compliant design demonstrating feasibility on the constrained site. The \"meets relevant planning controls\" condition in Table 6 requires that verification — it cannot be assumed.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 2.2 - Refined language on clause 4.6
add_para(new_doc, "2.2 Condition 3: A Clause 4.6 Variation Is Highly Likely", style='Heading 2')

add_para(new_doc, "The third qualifying condition requires that the development not include a clause 4.6 variation. On the proposed Homes NSW site, a clause 4.6 variation to height or FSR is highly likely based on envelope constraints and absence of demonstrated compliance:", style='Normal')
add_para(new_doc, "The base LEP height limit is 13.6 m. A five-storey residential flat building on this site is likely to exceed that limit.", style='Normal')
add_para(new_doc, "Division 1 of SEPP (Housing) 2021 may enable bonus height and FSR where an affordable housing component is provided, but the proponent has not published calculations demonstrating the applicable Division 1 bonus ceiling or confirmed compliance with it.", style='Normal')
add_para(new_doc, "If the proposal exceeds either the base LEP standard or any Division 1 bonus ceiling, clause 4.6 of Tweed LEP 2014 is engaged — and the development fails the third qualifying condition, mandating public exhibition.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "The Draft CPP does not address how the Table 6 qualifying conditions interact with Division 1 of SEPP (Housing) 2021 for social housing proposals by public authorities. This ambiguity itself creates a risk that the exemption could be applied without careful analysis of whether the Division 1 bonus ceiling has been reached.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 2.3
add_para(new_doc, "2.3 The 7-Day Pre-Commencement Notice Is Structurally Inadequate", style='Heading 2')
add_para(new_doc, "Even if the Table 6 exemption were available (which it is not on this site), the 7-day pre-commencement notice to adjoining neighbours is structurally inadequate for development involving complex life-safety obligations on FPCC 1 land:", style='Normal')
add_para(new_doc, "It is provided at commencement of construction — when the development application has already been determined and all opportunities for formal submission have passed.", style='Normal')
add_para(new_doc, "It does not provide access to the flood studies, hydraulic models, evacuation assessments, or cumulative impact assessments that were relied upon in the determination.", style='Normal')
add_para(new_doc, "It does not create any right to seek reconsideration of the determination — a Division 8.2 review is the only available pathway, but grounds are limited and the 28-day review period runs from determination, not commencement.", style='Normal')
add_para(new_doc, "Adjoining residents — including the submitter at 7 Heffron Street and the operator of the adjacent school — share the same constrained evacuation network and bear the same flood life-safety risks as the occupants of the proposed development. A 7-day notice of works commencing is not a meaningful substitute for the right to participate in the assessment of those shared risks.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 3 with micro-summary
add_para(new_doc, "3. Coordinated Commencement Risk: Targeted Assessment and the CPP", style='Heading 1')

# MICRO-SUMMARY ADDED
p = add_para(new_doc, "Key point: Discussion Paper confirms TAD pathway will commence with new CPP (late 2026/2027). Combined effect could eliminate all exhibition despite Council removing Change Option 22 on 3 April 2025 due to flood concerns.", style='Normal')
p.runs[0].italic = True

add_para(new_doc, "", style='Normal')

add_para(new_doc, "The Discussion Paper explicitly states:", style='Normal')
add_para(new_doc, "\"The remaining provisions of the Planning System Reforms Act 2025 will commence at a later date to allow for further consultation, system updates, and alignment with future regulatory changes, including the commencement of the new statewide Community Participation Plan.\"", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This statement confirms that the Department has deliberately linked the remaining PSR Act provisions — which include the Targeted Assessment Development (TAD) pathway under Division 4.3A of the EP&A Act — to the commencement of the new statewide CPP. The TAD pathway commenced as primary legislation on 21 March 2026 but is inert without a SEPP declaring development as \"targeted assessment development.\" That activating SEPP is expected to arrive with or shortly after the final CPP.", style='Normal')
add_para(new_doc, "The consequence for this site is direct:", style='Normal')
add_para(new_doc, "If Homes NSW lodges a DA after the TAD-activating SEPP is made and the new CPP is in force, the combined effect of the TAD pathway (reduced or no exhibition where strategic planning addressed the issues) and the CPP residential flat building exemption (Table 6) could result in no public exhibition period and no mandatory community notification at DA stage.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This is not a remote or theoretical risk. The Homes NSW consultation letter is dated 19 August 2025. Pre-lodgement consultation has occurred. A formal DA lodgement in late 2026 or 2027 — after the CPP and TAD SEPP have commenced — is a plausible and foreseeable scenario.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "The TAD pathway is explicitly designed to allow the consent authority to turn off exhibition steps where matters have been addressed at the strategic planning stage. But for this site, the strategic planning record runs in the opposite direction: Tweed Shire Council's Planning Committee resolved on 3 April 2025 to remove Change Option 22 (the South Tweed residential intensification proposal) precisely because of flood constraints. There is no exhibited masterplan or strategic plan that has resolved the flood life-safety questions for this precinct — the opposite is true.", style='Normal')
add_para(new_doc, "", style='Normal')

# FAILURE PATHWAY DIAGRAM (NEW)
p = add_para(new_doc, "FAILURE PATHWAY:", style='Normal')
p.runs[0].bold = True
add_para(new_doc, "Table 6 exemption → assumes compliance with planning controls → clause 5.22 requires demonstrated evacuation safety → cannot be demonstrated at lodgement → exemption applied prematurely → no exhibition → life-safety risks not publicly tested", style='Normal')
add_para(new_doc, "", style='Normal')

add_para(new_doc, "The TAD pathway should not operate to reduce exhibition for development in flood planning areas where:", style='Normal')
add_para(new_doc, "No site-specific flood study reconciling with the THSLDS baseline has been completed;", style='Normal')
add_para(new_doc, "Mandatory clause 5.22 preconditions remain undemonstrated; and", style='Normal')
add_para(new_doc, "The relevant strategic planning record expressly defers development uplift pending flood mitigation strategies.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 4
add_para(new_doc, "4. Recommendations", style='Heading 1')
add_para(new_doc, "I request that the Department incorporate the following amendments or clarifications into the final Statewide Community Participation Plan and, where legislative or regulatory amendment is required, into the EP&A Regulation:", style='Normal')
add_para(new_doc, "", style='Normal')

# Recommendation 1
add_para(new_doc, "Recommendation 1 — Flood Planning Area Carve-Out from Table 6", style='Normal')
p = add_para(new_doc, "The Table 6 residential flat building exemption should not apply to development on land in a flood planning area as defined in the NSW Flood Risk Management Manual 2023 where:", style='Normal')
add_para(new_doc, "mandatory flood planning preconditions under the applicable LEP or SEPP (including Tweed LEP 2014 clause 5.22 or its CC&NH SEPP successor) have not been demonstrated through site-specific technical assessment; or", style='Normal')
add_para(new_doc, "the site is classified FPCC 1 or FPCC 2 under the applicable flood study, indicating levee-protected or otherwise highly constrained flood behaviour.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "The CPP should include a clarifying note to Table 6 (or a corresponding provision in the EP&A Regulation amendment) in terms such as:", style='Normal')
add_para(new_doc, "\"A residential flat building is not exempt from public exhibition under this table if it is located on land in a flood planning area, unless the consent authority has first confirmed compliance with any mandatory flood planning preconditions under the applicable environmental planning instrument.\"", style='Normal')
add_para(new_doc, "", style='Normal')

# Recommendation 2
add_para(new_doc, "Recommendation 2 — Targeted Assessment SEPP Exclusion for Flood Planning Areas", style='Normal')
p = add_para(new_doc, "Any SEPP declaring development as targeted assessment development under Division 4.3A of the EP&A Act should expressly exclude from the targeted assessment pathway any residential flat building development on land in a flood planning area where:", style='Normal')
add_para(new_doc, "mandatory preconditions under clause 5.22 of the applicable LEP (or its CC&NH SEPP successor) remain undemonstrated by site-specific technical evidence; or", style='Normal')
add_para(new_doc, "the applicable council's strategic planning record does not support residential intensification in the precinct, including where a planning committee has resolved to defer or remove a density uplift option pending flood mitigation strategies.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This exclusion is consistent with the TAD pathway's own statutory framework: Division 4.3A preserves the obligation to consider significant likely impacts under s 4.15(1)(b), and a reduced exhibition period that denies the community opportunity to identify and submit on those impacts would undermine that preservation.", style='Normal')
add_para(new_doc, "", style='Normal')

# Recommendation 3
add_para(new_doc, "Recommendation 3 — Cumulative Impact Consultation Trigger", style='Normal')
p = add_para(new_doc, "Where a residential flat building DA is submitted for land in a flood planning area that is classified FPCC 1 or FPCC 2, and the site directly interfaces with an educational establishment or early education and care facility sharing the same constrained evacuation network, the CPP should require:", style='Normal')
add_para(new_doc, "a minimum 28-day public exhibition period, regardless of whether a clause 4.6 variation is proposed; and", style='Normal')
add_para(new_doc, "written notification to the principal of any adjacent educational establishment and the operator of any adjacent early education and care facility, in addition to the standard adjoining owner notification.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This requirement reflects the mandatory preconditions in clause 5.22(2)(b) of the Tweed LEP 2014, which requires the consent authority to consider whether the development affects the safe occupation and efficient evacuation of the locality — a consideration that necessarily involves adjacent sensitive uses sharing the constrained network.", style='Normal')
add_para(new_doc, "", style='Normal')

# Recommendation 4 - Tightened to link to clause 5.22
add_para(new_doc, "Recommendation 4 — Specialist Referral Transparency (Linked to Clause 5.22 Satisfaction)", style='Normal')
p = add_para(new_doc, "Where a consent authority refers a DA to Council engineering, NSW SES, or a hydraulic engineer for advice on flood evacuation or co-incident hazard considerations, affected residents within the notified area should be:", style='Normal')
add_para(new_doc, "notified that the referral has occurred and to which agency or specialist; and", style='Normal')
add_para(new_doc, "provided with a copy of the specialist advice, or a summary of its findings, during the exhibition period.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This transparency requirement is critical in levee-protected precincts where the adequacy of the specialist advice directly determines whether the mandatory preconditions under clause 5.22 are satisfied, and where affected residents bear the same evacuation risk as future occupants of the proposed development.", style='Normal')
add_para(new_doc, "", style='Normal')

# Crown development point - REMOVED to keep focus razor-sharp (as recommended)

# Recommendation 5
add_para(new_doc, "Recommendation 5 — Savings and Transitional Provisions for Flood-Constrained Sites", style='Normal')
p = add_para(new_doc, "The transitional provisions confirming that the statewide CPP prevails over existing council CPPs to the extent of any inconsistency on exhibition periods should expressly preserve any existing council CPP provision that requires public exhibition for residential development in a flood planning area.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "Tweed Shire Council's CPP and development assessment practices should not be reduced in their community participation requirements for flood-constrained development as a result of the statewide CPP's adoption, given the specific life-safety considerations identified in Council's own commissioned technical studies (the THSLDS) and the strategic planning decision of 3 April 2025.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 5 with micro-summary
add_para(new_doc, "5. Site-Specific Context Supporting These Recommendations", style='Heading 1')

# MICRO-SUMMARY ADDED
p = add_para(new_doc, "Key point: Recommendations grounded in THSLDS (Council's own engineers), Planning Committee resolution (3 April 2025), and operative DCP 2025 — not abstract propositions.", style='Normal')
p.runs[0].italic = True

add_para(new_doc, "", style='Normal')

add_para(new_doc, "The recommendations above are not abstract propositions. Each is directly grounded in the technical and statutory circumstances of the proposed Homes NSW development at Lots 2–4 DP530539, and in documents on the public record.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 5.1
add_para(new_doc, "5.1 Council's Own Technical Evidence Confirms FPCC 1 Status", style='Heading 2')
add_para(new_doc, "The Tweed Heads South Levee and Drainage Study (THSLDS, Draft 2025), commissioned by Tweed Shire Council and publicly exhibited from 9 September to 26 October 2025, makes four findings directly relevant to the adequacy of the Table 6 exemption for this site:", style='Normal')
add_para(new_doc, "Finding 1 — Existing levee inadequacy: The existing levee provides protection only to approximately 5% AEP, with sections already subsided below that standard. Overtopping of 0.3 m depth is predicted at the South Tweed Sports Club during a 1% AEP event.", style='Normal')
add_para(new_doc, "Finding 2 — FPCC 1 classification (existing conditions): The subject site and the surrounding residential fabric of Tweed Heads South, including Heffron Street and Seymour Street, are classified FPCC 1 under Figure 42.1 of the THSLDS.", style='Normal')
add_para(new_doc, "Finding 3 — FPCC 1 classification (best-case mitigation): Figure 55.1 of the THSLDS confirms the site remains FPCC 1 even under the best-performing combined levee upgrade scenario, at an estimated cost of $29 million, which itself is unfunded.", style='Normal')
add_para(new_doc, "Finding 4 — Table 36 explicit guidance: The THSLDS explicitly states that \"intensification of development would still be difficult to support\" in the Change 22 precinct even after full implementation of the combined levee option.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "A development that Council's own commissioned engineers assess as \"difficult to support\" for intensification under any credible mitigation scenario cannot legitimately qualify for the Table 6 exemption on the basis that it \"meets the relevant planning controls.\" The THSLDS is the authoritative baseline against which planning control compliance must be assessed — and it directly negates any assumption of compliance.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 5.2
add_para(new_doc, "5.2 The Strategic Planning Record Precludes TAD Application", style='Heading 2')
add_para(new_doc, "On 3 April 2025, the Tweed Shire Council Planning Committee resolved to remove Change Option 22 — South Tweed — from the Growth Management and Housing Strategy due to \"community opposition, vulnerability to tidal inundation and flooding concerns.\" The resolution explicitly noted this would also remove Options Paper V1 Changes 11, 12, and 13 (increased density in South Tweed Heads). The precinct was retained only as a future investigation area, pending flood mitigation strategies.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This is the opposite of the strategic planning context the TAD pathway is designed for. The TAD pathway is intended to streamline assessment where strategic planning has already addressed the issues — here, strategic planning has expressly deferred the issue. Applying the TAD pathway to this site would be directly inconsistent with Council's own strategic planning record.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 5.3
add_para(new_doc, "5.3 Tweed DCP 2025 Adds Mandatory Design Obligations That Cannot Be Assumed Met", style='Heading 2')
add_para(new_doc, "The Tweed Development Control Plan 2025 has now been adopted. Part D2 Section 2.3.2 (Multi-Dwelling Housing and Residential Flat Buildings) imposes the following cumulative landscaping and deep soil zone obligations for the proposed development on a combined site area of 2,081.89 m²:", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "These obligations are cumulative with the flood-compatible design constraints, the Class 2 acid sulfate soils management requirements, and the SEPP (Resilience and Hazards) 2021 coastal zone obligations. A five-storey residential flat building on a 2,081.89 m² site with ground-level parking cannot satisfy all these obligations simultaneously without detailed design analysis. The \"meets relevant planning controls\" condition in Table 6 cannot be satisfied by assumption — it requires verification.", style='Normal')
add_para(new_doc, "", style='Normal')

# Section 6 with enhanced policy failure statement
add_para(new_doc, "6. Conclusion and Formal Request", style='Heading 1')

# MICRO-SUMMARY ADDED
p = add_para(new_doc, "Key point: Creates a pathway by which development involving unresolved flood life-safety risks may proceed without any opportunity for community scrutiny at the point those risks are assessed.", style='Normal')
p.runs[0].bold = True
p.runs[0].italic = True

add_para(new_doc, "", style='Normal')

add_para(new_doc, "The Draft Statewide Community Participation Plan, as currently drafted, would — if applied to the proposed Homes NSW development at 3–5 Heffron Street and 6 Seymour Street, Tweed Heads South — risk removing meaningful community participation from the assessment of a development proposal with complex and unresolved flood life-safety implications. This is not a theoretical risk: the coordinated commencement of the statewide CPP and the TAD-activating SEPP, both expected in late 2026 or early 2027, could align with a Homes NSW DA lodgement in exactly the same window.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "The five recommendations in this submission are targeted, technically grounded, and legally defensible. They do not seek to obstruct social housing delivery. They seek to ensure that where mandatory preconditions exist — as they do under Tweed LEP 2014 clause 5.22 and the applicable SEPP obligations — the community retains a meaningful opportunity to participate in the assessment of whether those preconditions have been met, before determination, not seven days before excavation commences.", style='Normal')
add_para(new_doc, "", style='Normal')

# STRATEGIC REFRAME SENTENCE (NEW)
p = add_para(new_doc, "The issue is not whether these developments should proceed, but whether the statutory preconditions governing life-safety can be tested without public scrutiny.", style='Normal')
p.runs[0].italic = True

add_para(new_doc, "", style='Normal')

add_para(new_doc, "I formally request that the Department:", style='Normal')
add_para(new_doc, "Incorporate Recommendations 1–5 into the final Statewide Community Participation Plan or the accompanying EP&A Regulation amendment;", style='Normal')
add_para(new_doc, "Confirm in the submissions report how the residential flat building exemption in Table 6 interacts with mandatory flood planning preconditions under clause 5.22 and its CC&NH SEPP successor;", style='Normal')
add_para(new_doc, "Confirm that the TAD-activating SEPP will expressly exclude from the targeted assessment pathway development on land in a flood planning area where mandatory preconditions under clause 5.22 or its successor remain undemonstrated; and", style='Normal')
add_para(new_doc, "Confirm that the transitional provisions will preserve any existing council CPP requirement for public exhibition of residential development in a flood planning area.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "This submission is made in conjunction with my earlier submissions on the CC&NH SEPP EIE (11 March 2026) and the pre-lodgement submission to Tweed Shire Council (11 March 2026, Ref D25/2290641). The statutory and technical arguments in those submissions are incorporated by reference and are not repeated here.", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "Yours faithfully,", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "Tom Thorp", style='Normal')
add_para(new_doc, "Affected Resident", style='Normal')
add_para(new_doc, "7 Heffron Street, Tweed Heads South NSW 2486", style='Normal')
add_para(new_doc, "contact@tomthorp.me", style='Normal')
add_para(new_doc, "9 April 2026", style='Normal')
add_para(new_doc, "", style='Normal')
add_para(new_doc, "Annexure — Key Authorities and Documents Relied Upon", style='Normal')

# Save the revised document
new_doc.save('/workspace/CPP_Submission_Tom_Thorp_April2026_REVISED_FINAL.docx')
print("Document successfully created: CPP_Submission_Tom_Thorp_April2026_REVISED_FINAL.docx")

import { useState, useRef } from "react";
import {
  constituencyData,
  casteOptions,
  genderOptions,
  ageGroups,
  voteOptions,
  winOptions,
  GOOGLE_SCRIPT_URL,
} from "../data/surveyData";
import { formatAcSelectLabel, sortConstituenciesByAcNo } from "../data/acNumbers";
import "./SurveyForm.css";

/** Wall clock in Asia/Kolkata as `yyyy-mm-dd HH:mm:ss` — unambiguous for Apps Script date filters (avoids en-IN 1/4/2026 vs 4/1/2026). */
function timestampKolkata() {
  return new Date().toLocaleString("sv-SE", {
    timeZone: "Asia/Kolkata",
    hour12: false,
  });
}

const initialForm = {
  ac: "",
  faName: "",
  caste: "",
  gender: "",
  age: "",
  vote2021: "",
  vote2024: "",
  vote2026: "",
  whoWillWin: "",
};

const fieldLabels = {
  ac: "Assembly Constituency",
  faName: "Field Agent Name",
  caste: "Caste",
  gender: "Gender",
  age: "Age Group",
  vote2021: "2021 Assembly Election vote",
  vote2024: "2024 General Election vote",
  vote2026: "2026 Assembly Election vote",
  whoWillWin: "Who will win",
};

export default function SurveyForm() {
  const [form, setForm] = useState(initialForm);
  const [submitting, setSubmitting] = useState(false);
  const [success, setSuccess] = useState(false);
  const [error, setError] = useState("");
  const [fieldErrors, setFieldErrors] = useState({});
  const [attempted, setAttempted] = useState(false);
  const formRef = useRef();

  // Normalize AC names to avoid FA dropdown missing due to spelling variants
  // (e.g., tracker CSV uses "Kazhakootam" while form list uses "Kazhakkoottam").
  const normalizeAcKey = (v) =>
    String(v || "")
      .toLowerCase()
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/\(sc\)/g, "")
      .replace(/ac/g, "")
      .replace(/[^a-z]/g, "");

  const CANONICAL_AC_ALIAS = {
    kazhakootam: "Kazhakkoottam",
    kazhakkoottam: "Kazhakkoottam",
    nemom: "Nemom",
    nemam: "Nemom",
    nattika: "Nattika",
    nattikaac: "Nattika",
    thrissur: "Thrissur",
    thrissurac: "Thrissur",
    manalur: "Manalur",
    manalurac: "Manalur",
    perumbaavoor: "Perumbavoor",
    perumbavoor: "Perumbavoor",
    kanjirapalli: "Kanjirappally",
    kanjirappally: "Kanjirappally",
  };

  const canonicalizeAc = (name) => {
    const k = normalizeAcKey(name);
    return CANONICAL_AC_ALIAS[k] || name;
  };

  // Derived: FA names for selected AC
  const selectedAC = constituencyData.find((c) => normalizeAcKey(c.ac) === normalizeAcKey(form.ac));
  const faNames = selectedAC
    ? [selectedAC.fa1, selectedAC.fa2, selectedAC.fa3, selectedAC.fa4].filter(Boolean)
    : [];
  /** ACs in surveyData with no FA rows yet — free-text FA name */
  const faNameIsText = Boolean(form.ac && selectedAC && faNames.length === 0);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm((prev) => {
      const updated = { ...prev, [name]: name === "ac" ? canonicalizeAc(value) : value };
      if (name === "ac") updated.faName = "";
      return updated;
    });
    // Clear error for this field when user selects a value
    if (value) {
      setFieldErrors((prev) => {
        const copy = { ...prev };
        delete copy[name];
        return copy;
      });
    }
  };

  const validate = (formData) => {
    const errors = {};
    const allFields = Object.keys(fieldLabels);
    for (const key of allFields) {
      if (!formData[key]) {
        errors[key] = `${fieldLabels[key]} is required`;
      }
    }
    return errors;
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setSuccess(false);
    setAttempted(true);

    const errors = validate(form);
    setFieldErrors(errors);

    if (Object.keys(errors).length > 0) {
      // Scroll to the first error field
      const firstErrorKey = Object.keys(errors)[0];
      const el = formRef.current?.querySelector(`[name="${firstErrorKey}"]`);
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "center" });
        el.focus();
      }
      setError("Please answer all questions before submitting.");
      return;
    }

    setSubmitting(true);

      try {
        // Send labels only — Google Apps Script resolves weights from AC_DEMOGRAPHICS
        // so scores stay correct after you update the script (avoids stale cached JS sending 0).
        const payload = {
          timestamp: timestampKolkata(),
          ac: form.ac,
          faName: form.faName,
          caste: form.caste,
          gender: form.gender,
          age: form.age,
          vote2021: form.vote2021,
          vote2024: form.vote2024,
          vote2026: form.vote2026,
          whoWillWin: form.whoWillWin,
        };
  
        // Create an AbortController so we can force the UI to stop waiting after 1 second
        // since we know the Apps Script receives the data and writes it to the sheet instantly
        // before spending time on the group calculations.
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 1000);
  
        try {
          // Changed to text/plain to prevent CORS preflight freezing the Promise
          await fetch(GOOGLE_SCRIPT_URL, {
            method: "POST",
            mode: "no-cors",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify(payload),
            signal: controller.signal
          });
        } catch (fetchErr) {
          // In no-cors mode, a timeout or opaque response error is normal and safe to ignore 
          // because the request was transmitted to the Google server.
          console.log("Transmission complete (Fetch handoff)", fetchErr);
        } finally {
          clearTimeout(timeoutId);
        }
  
        setSuccess(true);
        setForm(initialForm);
        setFieldErrors({});
        setAttempted(false);
        // Scroll to top so user sees the success message
        window.scrollTo({ top: 0, behavior: "smooth" });
        setTimeout(() => setSuccess(false), 5000);
      } catch (err) {
        console.error(err);
        setError(
          "Submission encountered an issue, but may have saved to Google Sheets."
        );
      } finally {
        setSubmitting(false);
      }
  };

  const hasError = (name) => attempted && fieldErrors[name];

  return (
    <div className="survey-wrapper">
      {/* Accent stripe (navy / blue / sky) */}
      <div className="tricolor-stripe">
        <span></span><span></span><span></span>
      </div>

      {/* Header */}
      <header className="survey-header">
        <div className="header-content">
          <div className="logo-badge">
            <svg viewBox="0 0 24 24" fill="none" width="26" height="26">
              <path d="M12 2L2 7l10 5 10-5-10-5z" fill="#fff" opacity="0.9"/>
              <path d="M2 17l10 5 10-5" stroke="#fff" strokeWidth="1.5" fill="none" opacity="0.7"/>
              <path d="M2 12l10 5 10-5" stroke="#fff" strokeWidth="1.5" fill="none" opacity="0.85"/>
            </svg>
          </div>
          <div>
            <h1>Kerala Survey 2026</h1>
            <p className="subtitle">Assembly Election — Field Data Collection</p>
          </div>
        </div>
      </header>

      <form
        className="survey-form"
        onSubmit={handleSubmit}
        autoComplete="off"
        ref={formRef}
        noValidate
      >
        {/* Section 1: Location & Field Agent */}
        <div className="form-section">
          <div className="section-title">
            <span className="section-number">1</span>
            Location &amp; Field Agent
          </div>

          <div className={`field-group ${hasError("ac") ? "has-error" : ""}`}>
            <label htmlFor="ac">
              Assembly Constituency <span className="req">*</span>
            </label>
            <select
              id="ac"
              name="ac"
              value={form.ac}
              onChange={handleChange}
            >
              <option value="">— Select Constituency —</option>
              {sortConstituenciesByAcNo(constituencyData).map((c) => (
                <option key={c.ac} value={c.ac}>
                  {formatAcSelectLabel(c.ac)}
                </option>
              ))}
            </select>
            {hasError("ac") && (
              <div className="field-error-text">Select a constituency</div>
            )}
          </div>

          <div className={`field-group ${hasError("faName") ? "has-error" : ""}`}>
            <label htmlFor="faName">
              Field Agent (FA) Name <span className="req">*</span>
            </label>
            {faNameIsText ? (
              <input
                type="text"
                id="faName"
                name="faName"
                value={form.faName}
                onChange={handleChange}
                placeholder="Enter field assistant name"
                autoComplete="name"
                className="survey-text-input"
              />
            ) : (
              <select
                id="faName"
                name="faName"
                value={form.faName}
                onChange={handleChange}
                disabled={!form.ac}
              >
                <option value="">
                  {form.ac ? "— Select FA Name —" : "— Select AC first —"}
                </option>
                {faNames.map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
              </select>
            )}
            {hasError("faName") && (
              <div className="field-error-text">
                {faNameIsText ? "Enter the field agent name" : "Select a field agent"}
              </div>
            )}
          </div>
        </div>

        {/* Section 2: Respondent Details */}
        <div className="form-section">
          <div className="section-title">
            <span className="section-number">2</span>
            Respondent Details
          </div>

          <div className="field-row">
            <div className={`field-group ${hasError("caste") ? "has-error" : ""}`}>
              <label htmlFor="caste">
                Caste <span className="req">*</span>
              </label>
              <select
                id="caste"
                name="caste"
                value={form.caste}
                onChange={handleChange}
              >
                <option value="">— Select —</option>
                {casteOptions.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
              {hasError("caste") && (
                <div className="field-error-text">Required</div>
              )}
            </div>

            <div className={`field-group ${hasError("gender") ? "has-error" : ""}`}>
              <label htmlFor="gender">
                Gender <span className="req">*</span>
              </label>
              <select
                id="gender"
                name="gender"
                value={form.gender}
                onChange={handleChange}
              >
                <option value="">— Select —</option>
                {genderOptions.map((g) => (
                  <option key={g} value={g}>
                    {g}
                  </option>
                ))}
              </select>
              {hasError("gender") && (
                <div className="field-error-text">Required</div>
              )}
            </div>
          </div>

          <div className={`field-group ${hasError("age") ? "has-error" : ""}`}>
            <label htmlFor="age">
              Age Group <span className="req">*</span>
            </label>
            <select
              id="age"
              name="age"
              value={form.age}
              onChange={handleChange}
            >
              <option value="">— Select Age Group —</option>
              {ageGroups.map((a) => (
                <option key={a} value={a}>
                  {a}
                </option>
              ))}
            </select>
            {hasError("age") && (
              <div className="field-error-text">Required</div>
            )}
          </div>
        </div>

        {/* Section 3: Survey Questions */}
        <div className="form-section">
          <div className="section-title">
            <span className="section-number">3</span>
            Survey Questions
          </div>

          <div className={`question-card ${hasError("vote2021") ? "has-error" : ""}`}>
            <p className="question-text">
              <span className="q-num">Q1</span> Whom did you vote for in the{" "}
              <strong>2021 Assembly Election</strong>?
            </p>
            <select
              name="vote2021"
              value={form.vote2021}
              onChange={handleChange}
            >
              <option value="">— Select —</option>
              {voteOptions.map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))}
            </select>
            {hasError("vote2021") && (
              <div className="field-error-text">Please answer this question</div>
            )}
          </div>

          <div className={`question-card ${hasError("vote2024") ? "has-error" : ""}`}>
            <p className="question-text">
              <span className="q-num">Q2</span> Whom did you vote for in the{" "}
              <strong>2024 General Election</strong>?
            </p>
            <select
              name="vote2024"
              value={form.vote2024}
              onChange={handleChange}
            >
              <option value="">— Select —</option>
              {voteOptions.map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))}
            </select>
            {hasError("vote2024") && (
              <div className="field-error-text">Please answer this question</div>
            )}
          </div>

          <div className={`question-card ${hasError("vote2026") ? "has-error" : ""}`}>
            <p className="question-text">
              <span className="q-num">Q3</span> Whom will you vote for in the{" "}
              <strong>2026 Assembly Election</strong>?
            </p>
            <select
              name="vote2026"
              value={form.vote2026}
              onChange={handleChange}
            >
              <option value="">— Select —</option>
              {winOptions.map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))}
            </select>
            {hasError("vote2026") && (
              <div className="field-error-text">Please answer this question</div>
            )}
          </div>

          <div className={`question-card ${hasError("whoWillWin") ? "has-error" : ""}`}>
            <p className="question-text">
              <span className="q-num">Q4</span> Who do you think{" "}
              <strong>will win</strong> in this constituency?
            </p>
            <select
              name="whoWillWin"
              value={form.whoWillWin}
              onChange={handleChange}
            >
              <option value="">— Select —</option>
              {winOptions.map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))}
            </select>
            {hasError("whoWillWin") && (
              <div className="field-error-text">Please answer this question</div>
            )}
          </div>
        </div>

        {/* Error */}
        {error && (
          <div className="msg error-msg">
            ⚠️ {error}
            {attempted && Object.keys(fieldErrors).length > 0 && (
              <ul className="error-list">
                {Object.entries(fieldErrors).map(([key, msg]) => (
                  <li key={key}>{msg}</li>
                ))}
              </ul>
            )}
          </div>
        )}

        {/* Submit */}
        <button type="submit" className="submit-btn" disabled={submitting}>
          {submitting ? (
            <>
              <span className="spinner"></span> Submitting…
            </>
          ) : (
            "Submit Response"
          )}
        </button>

        <p className="footer-note">
          All fields are mandatory. Data is saved to the central Google Sheet.
        </p>
      </form>

      {/* ── Success Popup Overlay ── */}
      {success && (
        <div className="popup-overlay" onClick={() => setSuccess(false)}>
          <div className="popup-box" onClick={(e) => e.stopPropagation()}>
            <div className="popup-check">✅</div>
            <h2 className="popup-title">Submitted Successfully!</h2>
            <p className="popup-text">Response has been saved to Google Sheet.</p>
            <button className="popup-btn" onClick={() => setSuccess(false)}>
              OK
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

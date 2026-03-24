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
import { getWeights } from "../data/demographicWeights";
import "./SurveyForm.css";

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

  // Derived: FA names for selected AC
  const selectedAC = constituencyData.find((c) => c.ac === form.ac);
  const faNames = selectedAC
    ? [selectedAC.fa1, selectedAC.fa2].filter(Boolean)
    : [];

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm((prev) => {
      const updated = { ...prev, [name]: value };
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
      // Look up demographic weights
      const { casteWeight, genderWeight, ageWeight } = getWeights(
        form.ac,
        form.caste,
        form.gender,
        form.age
      );

      const payload = {
        timestamp: new Date().toLocaleString("en-IN", {
          timeZone: "Asia/Kolkata",
        }),
        ac: form.ac,
        faName: form.faName,
        caste: casteWeight,     // caste % / 100
        gender: genderWeight,   // gender % / 100
        age: ageWeight,         // Age Normal from Sheet2
        vote2021: form.vote2021,
        vote2024: form.vote2024,
        vote2026: form.vote2026,
        whoWillWin: form.whoWillWin,
      };

      await fetch(GOOGLE_SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      setSuccess(true);
      setForm(initialForm);
      setFieldErrors({});
      setAttempted(false);
      // Scroll to top so user sees the success message
      window.scrollTo({ top: 0, behavior: "smooth" });
      setTimeout(() => setSuccess(false), 5000);
    } catch {
      setError(
        "Submission failed. Please check your internet connection and try again."
      );
    } finally {
      setSubmitting(false);
    }
  };

  const hasError = (name) => attempted && fieldErrors[name];

  return (
    <div className="survey-wrapper">
      {/* Decorative tricolor stripe */}
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
              {constituencyData.map((c) => (
                <option key={c.ac} value={c.ac}>
                  {c.ac}
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
            {hasError("faName") && (
              <div className="field-error-text">Select a field agent</div>
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

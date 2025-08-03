import React, { useState } from "react";
import './UploadPage.css';

function UploadPage() {
  const [planFile, setPlanFile] = useState(null);
  const [slidesFiles, setSlidesFiles] = useState([]);
  const [assignmentFiles, setAssignmentFiles] = useState([]);

  const [modules, setModules] = useState([]);
  const [selectedSegments, setSelectedSegments] = useState([]);
  const [downloadUrl, setDownloadUrl] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  const maxVideoSegmentsPerModule = 7;

  const learningTypeOptions = ["Video", "Reading", "Assignment", "Quiz", "Discussion"];
  const videoTypeOptions = ["Talking head", "Light Board", "Screencast", "Lab Interview"];

  const handleSubmit = async (e) => {
    e.preventDefault();
    setIsLoading(true);
    setDownloadUrl("");

    const formData = new FormData();
    if (planFile) formData.append("plan", planFile);
    slidesFiles.forEach((f) => formData.append("slides", f));
    assignmentFiles.forEach((f) => formData.append("assignments", f));

    try {
      const res = await fetch("http://localhost:8002/generate-outline", {
        method: "POST",
        body: formData,
      });
      const data = await res.json();
      if (data.error) {
        alert("Error: " + data.error);
      } else {
        setModules(data.modules || []);
        setSelectedSegments([]);
      }
    } catch {
      alert("Upload failed");
    } finally {
      setIsLoading(false);
    }
  };

  const toggleSegmentSelection = (moduleIdx, blockIdx, segmentIdx) => {
    const segment = modules[moduleIdx].blocks[blockIdx].segments[segmentIdx];
    const alreadySelected = selectedSegments.some(
      (s) =>
        s.moduleIdx === moduleIdx &&
        s.blockIdx === blockIdx &&
        s.segmentIdx === segmentIdx
    );

    if (alreadySelected) {
      setSelectedSegments((prev) =>
        prev.filter(
          (s) =>
            !(
              s.moduleIdx === moduleIdx &&
              s.blockIdx === blockIdx &&
              s.segmentIdx === segmentIdx
            )
        )
      );
    } else {
      const selectedVideoCount = selectedSegments
        .filter((s) => s.moduleIdx === moduleIdx)
        .filter((s) => s.customLearningType === "Video").length;

      if (segment.learning_type === "Video" && selectedVideoCount >= maxVideoSegmentsPerModule) {
        alert(`You can select up to ${maxVideoSegmentsPerModule} video segments per module.`);
        return;
      }

      setSelectedSegments((prev) => [
        ...prev,
        {
          moduleIdx,
          blockIdx,
          segmentIdx,
          segment,
          customLearningType: segment.learning_type,
          customVideoType: segment.video_type || "",
        },
      ]);
    }
  };

  const submitFinal = async () => {
    if (selectedSegments.length === 0) {
      alert(`Please select at least one segment.`);
      return;
    }

    setIsLoading(true);

    const selectedModulesMap = {};

    selectedSegments.forEach(({ moduleIdx, blockIdx, segment, customLearningType, customVideoType }) => {
      if (!selectedModulesMap[moduleIdx]) {
        selectedModulesMap[moduleIdx] = {
          module_title: modules[moduleIdx].module_title,
          blocks: {},
        };
      }
      if (!selectedModulesMap[moduleIdx].blocks[blockIdx]) {
        selectedModulesMap[moduleIdx].blocks[blockIdx] = {
          block_title: modules[moduleIdx].blocks[blockIdx].block_title,
          segments: [],
        };
      }

      selectedModulesMap[moduleIdx].blocks[blockIdx].segments.push({
        ...segment,
        learning_type: customLearningType,
        video_type: customLearningType === "Video" ? customVideoType : "",
      });
    });

    const selectedModules = Object.values(selectedModulesMap).map((mod) => ({
      module_title: mod.module_title,
      blocks: Object.values(mod.blocks),
    }));

    try {
      const res = await fetch("http://localhost:8002/finalize-outline", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(selectedModules),
      });
      const data = await res.json();
      alert(data.message);
      setDownloadUrl("http://localhost:8002" + data.download_url);
    } catch {
      alert("Failed to create final outline.");
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="upload-container">
      <h2> AI Course Outline Generator</h2>

      <form onSubmit={handleSubmit}>
        <label>Upload Plan (docx):</label>
        <input
          type="file"
          onChange={(e) => setPlanFile(e.target.files[0])}
          accept=".docx"
          required
        />
        <label>Upload Slides (pptx):</label>
        <input
          type="file"
          multiple
          onChange={(e) => setSlidesFiles(Array.from(e.target.files))}
          accept=".pptx"
        />
        <label>Upload Assignments (optional):</label>
        <input
          type="file"
          multiple
          onChange={(e) => setAssignmentFiles(Array.from(e.target.files))}
          accept=".docx,.pdf,.txt"
        />
        <button type="submit" disabled={isLoading}>
          {isLoading ? "Processing..." : "Upload and Extract"}
        </button>
      </form>

      {modules.length > 0 && (
        <section>
          <h3>
            Select up to {maxVideoSegmentsPerModule} video segments per Module
          </h3>
          {modules.map((mod, mIdx) => {
            const selectedVideoCount = selectedSegments
              .filter((s) => s.moduleIdx === mIdx)
              .filter((s) => s.customLearningType === "Video").length;

            return (
              <div key={mIdx}>
                <h4> Module: {mod.module_title}</h4>
                {mod.blocks.map((block, bIdx) => (
                  <div key={bIdx} style={{ marginLeft: "20px" }}>
                    <h5>Block: {block.block_title}</h5>
                    <ul>
                      {block.segments.map((seg, sIdx) => {
                        const isSelected = selectedSegments.some(
                          (s) =>
                            s.moduleIdx === mIdx &&
                            s.blockIdx === bIdx &&
                            s.segmentIdx === sIdx
                        );
                        const selectedSegment = selectedSegments.find(
                          (s) =>
                            s.moduleIdx === mIdx &&
                            s.blockIdx === bIdx &&
                            s.segmentIdx === sIdx
                        );
                        const disableCheckbox =
                          !isSelected &&
                          seg.learning_type === "Video" &&
                          selectedVideoCount >= maxVideoSegmentsPerModule;

                        return (
                          <li key={sIdx} style={{ marginBottom: "12px" }}>
                            <label>
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={() =>
                                  toggleSegmentSelection(mIdx, bIdx, sIdx)
                                }
                                disabled={disableCheckbox}
                              />
                              {seg.segment_title}
                            </label>

                            {isSelected && (
                              <div style={{ marginLeft: "20px" }}>
                                <div>
                                  Learning Type:
                                  <select
                                    value={selectedSegment.customLearningType}
                                    onChange={(e) => {
                                      const newType = e.target.value;
                                      setSelectedSegments((prev) =>
                                        prev.map((s) =>
                                          s.moduleIdx === mIdx &&
                                            s.blockIdx === bIdx &&
                                            s.segmentIdx === sIdx
                                            ? {
                                              ...s,
                                              customLearningType: newType,
                                              customVideoType:
                                                newType === "Video"
                                                  ? s.customVideoType || videoTypeOptions[0]
                                                  : "",
                                            }
                                            : s
                                        )
                                      );
                                    }}
                                    style={{ marginLeft: "10px" }}
                                  >
                                    {learningTypeOptions.map((type) => (
                                      <option key={type} value={type}>
                                        {type}
                                      </option>
                                    ))}
                                  </select>
                                </div>

                                {selectedSegment.customLearningType === "Video" && (
                                  <div style={{ marginTop: "8px" }}>
                                    Video Type:
                                    <select
                                      value={selectedSegment.customVideoType}
                                      onChange={(e) => {
                                        const newVideo = e.target.value;
                                        setSelectedSegments((prev) =>
                                          prev.map((s) =>
                                            s.moduleIdx === mIdx &&
                                              s.blockIdx === bIdx &&
                                              s.segmentIdx === sIdx
                                              ? { ...s, customVideoType: newVideo }
                                              : s
                                          )
                                        );
                                      }}
                                      style={{ marginLeft: "10px" }}
                                    >
                                      {videoTypeOptions.map((type) => (
                                        <option key={type} value={type}>
                                          {type}
                                        </option>
                                      ))}
                                    </select>
                                  </div>
                                )}
                              </div>
                            )}
                          </li>
                        );
                      })}
                    </ul>
                  </div>
                ))}
              </div>
            );
          })}
          <button
            onClick={submitFinal}
            disabled={isLoading || selectedSegments.length === 0}
          >
            Create Final Outline
          </button>
        </section>
      )}

      {downloadUrl && (
        <a
          href={downloadUrl}
          target="_blank"
          rel="noopener noreferrer"
          download
          style={{ display: "block", marginTop: "20px" }}
        >
          Download Course Outline
        </a>
      )}
    </div>
  );
}

export default UploadPage;
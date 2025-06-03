// public/app.js

// === Access Code ===
document.getElementById("unlockBtn").addEventListener("click", () => {
  const code = document.getElementById("accessCode").value.trim();
  const validCode = "letmein"; // Change this to your desired access code

  if (code === validCode) {
    localStorage.setItem("access_granted", "yes");
    document.getElementById("auth").style.display = "none";
    document.getElementById("app").style.display = "block";
  } else {
    alert("Invalid access code");
  }
});

window.addEventListener("load", () => {
  if (localStorage.getItem("access_granted") === "yes") {
    document.getElementById("auth").style.display = "none";
    document.getElementById("app").style.display = "block";
  }
});

// === Drag-and-Drop ===
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileUpload");

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  fileInput.files = e.dataTransfer.files;
});

// === Process Button ===
document.getElementById("processBtn").addEventListener("click", async () => {
  const files = fileInput.files;
  const progress = document.getElementById("progress");
  const downloadLinks = document.getElementById("downloadLinks");

  if (!files.length) {
    return alert("Please upload at least one file.");
  }

  downloadLinks.innerHTML = "";
  progress.innerText = "Processing files...";

  for (let i = 0; i < files.length; i++) {
    const formData = new FormData();
    formData.append("file", files[i]);

    try {
      const response = await fetch("/api/processFile", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        // Read the server’s error message and show it
        const errorText = await response.text();
        throw new Error(
          `Error processing "${files[i].name}": ${errorText}`
        );
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `Processed_${files[i].name}`;
      link.className = "download-link";
      link.textContent = `Download: ${files[i].name}`;
      downloadLinks.appendChild(link);
    } catch (err) {
      console.error(err);
      // Show a clear error message alongside the file name
      progress.innerText = `Error: ${err.message}`;
      return; // Stop processing further files
    }
  }

  progress.innerText = "Processing complete.";
});

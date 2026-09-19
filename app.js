// Replace this string with the Web App URL generated when you deploy Code.gs
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyV6NrtXW6cpl88etibvaHGO86UR0xP3YJ2G6MEtcbtxjO5FfvkO9vCac9SIfAuXxWc/exec";

document.getElementById("membershipForm").addEventListener("submit", async function(e) {
  e.preventDefault();
  
  const submitBtn = document.getElementById("submitBtn");
  submitBtn.disabled = true;
  submitBtn.innerText = "Submitting Data...";

  try {
    const fileToBase64 = (file) => new Promise((resolve, reject) => {
      if (!file) resolve(null);
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => resolve({
        base64: reader.result.split(',')[1],
        mimeType: file.type,
        fileName: file.name
      });
      reader.onerror = error => reject(error);
    });

    const nidFile = document.getElementById("nidImage").files[0];
    const photoFile = document.getElementById("profilePic").files[0];
    const paymentFile = document.getElementById("paymentProof").files[0];

    const payload = {
      nameBn: document.getElementById("nameBn").value,
      nameEn: document.getElementById("nameEn").value,
      fatherName: document.getElementById("fatherName").value,
      motherName: document.getElementById("motherName").value,
      dob: document.getElementById("dob").value,
      nationality: document.getElementById("nationality").value,
      bloodGroup: document.getElementById("bloodGroup").value,
      sscYear: document.getElementById("sscYear").value,
      studyDuration: document.getElementById("studyDuration").value,
      occupation: document.getElementById("occupation").value,
      organisation: document.getElementById("organisation").value,
      contactNo: document.getElementById("contactNo").value,
      email: document.getElementById("email").value,
      presentAddress: document.getElementById("presentAddress").value,
      permanentAddress: document.getElementById("permanentAddress").value,
      nidNumber: document.getElementById("nidNumber").value,
      membershipType: document.getElementById("membershipType").value,
      membershipFee: document.getElementById("membershipFee").value,
      paymentMethod: document.getElementById("paymentMethod").value,
      declaration: document.getElementById("declaration").checked,
      files: {
        nidImage: await fileToBase64(nidFile),
        profilePic: await fileToBase64(photoFile),
        paymentProof: await fileToBase64(paymentFile)
      }
    };

    const response = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      body: JSON.stringify(payload)
    });

    const result = await response.json();

    if (result.result === "success") {
      alert(`Registration Successful! Your ID is: ${result.registrationId}`);
      document.getElementById("membershipForm").reset();
    } else {
      alert("Error submitting form: " + result.message);
    }
  } catch (err) {
    alert("Submission failed: " + err.message);
  } finally {
    submitBtn.disabled = false;
    submitBtn.innerText = "Submit Application";
  }
});
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("get-slide-btn").onclick = () => {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const slides = asyncResult.value.slides;
            if (slides.length > 0) {
              const slideIndex = slides[0].index;
              document.getElementById("status").innerText = "Current Slide Number: " + slideIndex;
            } else {
              document.getElementById("status").innerText = "No slide selected.";
            }
          } else {
            document.getElementById("status").innerText = "Error: " + asyncResult.error.message;
          }
        }
      );
    };
  }
});

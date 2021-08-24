var canvas = document.getElementById("signature-pad");
// Adjust canvas coordinate space taking into account pixel ratio,
// to make it look crisp on mobile devices.
// This also causes canvas to be cleared.
function resizeCanvas() {
  // When zoomed out to less than 100%, for some very strange reason,
  // some browsers report devicePixelRatio as less than 1
  // and only part of the canvas is cleared then.
  var ratio = Math.max(window.devicePixelRatio || 1, 1);
  canvas.width = canvas.offsetWidth * ratio;
  canvas.height = canvas.offsetHeight * ratio;
  canvas.getContext("2d").scale(ratio, ratio);
}

window.onresize = resizeCanvas;
resizeCanvas();

var signaturePad = new SignaturePad(canvas);

document.getElementById("clear").addEventListener("click", function () {
  signaturePad.clear();
});

document.getElementById("draw").addEventListener("click", function () {
  var ctx = canvas.getContext("2d");
  console.log(ctx.globalCompositeOperation);
  ctx.globalCompositeOperation = "source-over"; // default value
});

document.getElementById("erase").addEventListener("click", function () {
  var ctx = canvas.getContext("2d");
  ctx.globalCompositeOperation = "destination-out";
});

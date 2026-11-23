// Close the language menu when someone clicks elsewhere or presses Escape.
// The menu itself is a plain <details> element and works without this script.
(function () {
  var menu = document.querySelector(".lang-switch details");
  if (!menu) return;
  document.addEventListener("click", function (e) {
    if (menu.open && !menu.contains(e.target)) menu.open = false;
  });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && menu.open) {
      menu.open = false;
      menu.querySelector("summary").focus();
    }
  });
})();

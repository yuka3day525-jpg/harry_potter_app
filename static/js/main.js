const form = document.querySelector("#spell-form");
const loading = document.querySelector("#loading");
const button = document.querySelector("#spell-button");

form.addEventListener("submit", function () {

    loading.style.display = "block";

    button.disabled = true;

    button.textContent = "魔法発動中...";

});
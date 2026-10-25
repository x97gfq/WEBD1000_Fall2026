// Shared by flex_playground.html and grid_playground.html.
// Each <fieldset data-selector="..."> picks some boxes on the stage, and each
// control inside it (data-prop="...") sets that CSS property on them.
// The code box shows the CSS you'd write: a property is left out while it
// still has its default value (data-default).
// (JavaScript comes later in the program. You don't need to read this to use the playground.)

const form = document.querySelector(".controls");
const container = document.querySelector(".container");
const code = document.querySelector(".code code");
const startCount = container.children.length;

function update() {
	let css = "";
	for (const item of container.children) {
		item.removeAttribute("style");
	}
	form.querySelectorAll("fieldset").forEach((set) => {
		const selector = set.dataset.selector;
		const targets = container.matches(selector) ? [container] : container.querySelectorAll(":scope > " + selector);
		let lines = "";
		set.querySelectorAll("[data-prop]").forEach((control) => {
			const prop = control.dataset.prop;
			targets.forEach((t) => t.style.setProperty(prop, control.value));
			if (control.value !== control.dataset.default) {
				lines += "   " + prop + ": " + control.value + ";\n";
			}
		});
		if (lines) {
			css += selector + " {\n" + lines + "}\n\n";
		}
	});
	code.textContent = css.trim() || "/* Everything is on its default */";
}

// Add or remove a box (never fewer than two)
function addItem() {
	const item = document.createElement("div");
	item.className = "item";
	item.textContent = container.children.length + 1;
	container.append(item);
	update();
}

function removeItem() {
	if (container.children.length > 2) {
		container.lastElementChild.remove();
	}
}

form.addEventListener("input", update);
form.querySelector(".add").addEventListener("click", addItem);
form.querySelector(".remove").addEventListener("click", removeItem);

// After a reset, put the starting number of boxes back
form.addEventListener("reset", () => {
	setTimeout(() => {
		while (container.children.length > startCount) removeItem();
		while (container.children.length < startCount) addItem();
		update();
	});
});

update();

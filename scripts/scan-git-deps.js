const fs = require("fs");
const path = require("path");

const patterns = [
  "ssh://git@github.com",
  "git@github.com:",
  "backend/.env.git",
];

const filesToScan = [
  "package-lock.json",
  path.join("backend", "package-lock.json"),
  path.join("server", "package-lock.json"),
  "npm-shrinkwrap.json",
  "yarn.lock",
  "pnpm-lock.yaml",
];

let found = false;

filesToScan.forEach((file) => {
  if (!fs.existsSync(file)) return;
  const content = fs.readFileSync(file, "utf8");
  patterns.forEach((pattern) => {
    if (content.includes(pattern)) {
      console.error(`Found forbidden git dependency pattern "${pattern}" in ${file}`);
      found = true;
    }
  });
});

if (found) {
  process.exit(1);
}

console.log("No forbidden git dependency patterns found.");

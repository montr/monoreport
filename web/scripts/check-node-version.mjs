const [major, minor] = process.versions.node.split(".").map(Number);

if (major < 22 || (major === 22 && minor < 15)) {
  console.error(`MonoReport UI requires Node.js 22.15 or newer; found ${process.versions.node}.`);
  process.exit(1);
}

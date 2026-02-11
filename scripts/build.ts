console.log("Building...");

console.log("NODE_ENV:", process.env.NODE_ENV);
console.log("BUN_PUBLIC_CLIENT_ID:", process.env.BUN_PUBLIC_CLIENT_ID);


const htmlFiles = Array.from(
    new Bun.Glob("./src/**/*.html").scanSync()
  );
  

const result = await Bun.build({
    entrypoints: htmlFiles,
    outdir: './dist',
    sourcemap: true,
    target: 'browser',
    minify: true,
    env: "BUN_PUBLIC_*",
    splitting: false
  })


if (!result.success) {
    console.error("Build failed:");
    for (const log of result.logs) {
        console.error(log);
    }
    process.exit(1);
}

console.log("Build succeeded");
console.log("Outputs:");

for (const output of result.outputs) {
    console.log("Output:", output.path);
}

if (result.logs.length > 0) {
    console.warn("Build succeeded with warnings:");
    for (const message of result.logs) {
        // Bun will pretty print the message object
        console.warn(message);
    }
}

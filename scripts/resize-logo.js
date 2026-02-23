const { execSync } = require("child_process");
const path = require("path");

const root = path.resolve(__dirname, "..");
const src = path.join(root, "Brand", "logo.png");
const sizes = [192, 512];

function runResize(size) {
  const dest = path.join(root, "Brand", `logo-${size}.png`);
  const ps = `
Add-Type -AssemblyName System.Drawing;
$img = [System.Drawing.Image]::FromFile('${src}');
$bmp = New-Object System.Drawing.Bitmap ${size}, ${size};
$g = [System.Drawing.Graphics]::FromImage($bmp);
$g.InterpolationMode = 'HighQualityBicubic';
$g.DrawImage($img, 0, 0, ${size}, ${size});
$bmp.Save('${dest}', [System.Drawing.Imaging.ImageFormat]::Png);
$g.Dispose(); $img.Dispose(); $bmp.Dispose();
`;
  execSync(`powershell -NoProfile -Command "${ps.replace(/\r?\n/g, " ")}"`, { stdio: "inherit" });
}

sizes.forEach(runResize);
console.log("Generated resized logos.");

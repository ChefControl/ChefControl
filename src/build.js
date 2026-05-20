#!/usr/bin/env node
// Build entrypoint: read data/resume.yaml, render dist/Oran_Dan_Resume.docx.

const fs = require("fs");
const path = require("path");
const yaml = require("js-yaml");
const { Packer } = require("docx");

const {
  sectionHeading,
  entryLine1,
  entryLine2,
  bullet,
  renderRuns,
  headerName,
  headerContact,
  skillLine,
  buildDocument,
} = require("./style-modern");

const ROOT = path.resolve(__dirname, "..");
const DATA = path.join(ROOT, "data", "resume.yaml");
const DIST = path.join(ROOT, "dist");
const OUT  = path.join(DIST, "Oran_Dan_Resume.docx");

const data = yaml.load(fs.readFileSync(DATA, "utf8"));

const body = [
  headerName(data.header.name_first, data.header.name_last),
  headerContact(data.header.contact, data.header.latest_release_url),

  sectionHeading("Experience"),
  ...data.experience.flatMap((job) => [
    entryLine1(job.company, job.location),
    entryLine2(job.title, job.dates),
    ...job.bullets.map((b) => bullet(renderRuns(b.runs))),
  ]),

  sectionHeading("Skills"),
  ...data.skills.map((s) => skillLine(s.label, s.value)),
];

const doc = buildDocument(body);

fs.mkdirSync(DIST, { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log(`Wrote: ${OUT} (${buf.length} bytes)`);
});

## spfx-one-drive-form

This is where you include your WebPart documentation.

### Online Work Bench

You can use an online version of the workbench by adding `_layouts/workbench.aspx` to the end of a site URL. This is needed so that we can develop against the SharePoint Online Rest API.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO

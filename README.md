## spfx-task-list-web-part

This is where you include your WebPart documentation.

### Building the code

```bash
1. Clone the repository
2. run npm install
3. run npm i @microsoft/sp-http@v1.4.1
4. run npm dedupe
5. And finally, run gulp-serve
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

Screen-Shot of final product from my workbench - custom Task List web part

![Final Screen of custom web part for task list](https://github.com/devoooo4u/spfx-task-list-web-part/blob/master/src/webparts/images/Capture.PNG?raw=true "Custom web part -task list")

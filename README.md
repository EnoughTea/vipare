# vipare

vipare is a simple command-line utility used to export pages from Visio diagrams. As you can guess, it uses COM interop with a hidden Visio instance, so Visio needs to be installed.

Since a diagram can contain utility pages which you don't want to export, this tool ignores pages with names starting with one of the following 3 symbols: `` ` `` `~` `!`.


### Usage examples

Export pages from one Visio diagram to png files and store them in the current directory:

    vipare "file 1.vsdx"

Export pages from three Visio diagrams to bmp files and store them in the specified output folder:

    vipare -f bmp -o "D:\Resulting images\" "file 1.vsdx" "..\another file 2.vsdx" "C:\some folder\other file 3.vsdx"


### Options

    -o, --output    (Default: <current directory>) Defines output directory for exported images.

    -f, --format    (Default: png) Format indicates which export filter to use.
                    Supply one of file formats supported by Visio export
                    (bmp, dib, dwg, dxf, emf, emz, gif, htm, jpg, png, svg, svgz, tif, or wmf).
                    Default preference settings for the specified filter will be used.

    --help          Display the help screen.


### NuGet references

You may notice that NuGet packages are not in the repository, so do not forget to set up package restoration in Visual Studio:

Tools menu → Options → Package Manager → General → "Allow NuGet to download missing packages during build" should be selected.

If you have a build server then it needs to be setup with an environment variable 'EnableNuGetPackageRestore' set to true.

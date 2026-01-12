{ pkgs }: {
  deps = [
    pkgs.python310
    pkgs.tesseract4
    pkgs.poppler_utils
    pkgs.libGL
    pkgs.pango
    pkgs.cairo
    pkgs.gdk-pixbuf
  ];
}

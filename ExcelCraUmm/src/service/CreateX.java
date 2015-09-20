package service;

public class CreateX {

	private String rootpath;
	private String filename;
	public String getRootpath() {
		String root=rootpath.replace("\\", "\\\\");
		return root;
	}
	public void setRootpath(String rootpath) {
		this.rootpath = rootpath;
	}
	public String getFilename() {
		return filename;
	}
	public void setFilename(String filename) {
		this.filename = filename;
	}
}

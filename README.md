To install the Argentum Online WorldEditor and set it up correctly, you'll need to follow a few steps. These instructions involve downloading necessary files from GitHub and organizing them on your local system. Here's a simplified and clear guide to do just that:

### Step 1: Download the Resources
First, you need to download the game's resources. These resources are crucial for the WorldEditor to function properly as they contain the assets needed for map creation and editing. Open your terminal or command prompt and execute the following command to clone the resources repository from GitHub:

```bash
git clone https://github.com/ao-org/Recursos.git
```

This command will create a directory named `Recursos` in your current working directory and download the assets into this folder.

### Step 2: Download the WorldEditor
Next, you need to obtain the WorldEditor itself. You can do this in two ways: by cloning the repository if you are familiar with Git or by downloading the latest release from the project's GitHub releases page.

#### Option A: Clone the WorldEditor Repository
If you prefer to use Git, open your terminal or command prompt and run:

```bash
git clone https://github.com/ao-org/argentum-online-worldeditor.git
```

This will create a directory named `argentum-online-worldeditor` and download the WorldEditor's files.

#### Option B: Download the Latest Release
If you're not comfortable with Git or prefer using a graphical interface, follow these steps:

1. Visit the releases page of the WorldEditor at [https://github.com/ao-org/argentum-online-worldeditor/releases](https://github.com/ao-org/argentum-online-worldeditor/releases).
2. Download the latest release to your computer.
3. Extract the downloaded archive to a location of your choice.

### Step 3: Organize the Folders
After downloading both the `Recursos` and the WorldEditor (either through cloning or downloading the release), you need to ensure that both directories are at the same level in your filesystem. This means they should be located within the same parent directory, but not inside one another. Here's an example of a correct structure:

```
some_folder/
├── Recursos/
└── argentum-online-worldeditor/
```

This structure is necessary because the WorldEditor expects the resources to be accessible in a specific relative path to operate correctly.

### Step 4: Start Using the WorldEditor
With the resources and the WorldEditor downloaded and properly organized, you can now navigate to the `argentum-online-worldeditor` directory and start using the WorldEditor as per its documentation.

Following these steps will ensure that you have the WorldEditor set up correctly with all necessary resources. If you encounter any issues, consult the documentation or seek help from the Argentum Online community.

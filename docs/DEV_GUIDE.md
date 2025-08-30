# Developer Guide

Follow this guide if you want to build Lokka from source to contribute to the project.

## Pre-requisites

- Follow the [installation guide](../website/docs/install.mdx) to install Node.js
- Follow the [advanced guide](../website/docs/install-advanced/readme.md) if you wish to create a custom Entra application
- Clone the Lokka repository from GitHub: https://github.com/merill/lokka

## Building the Project

1. Open a terminal and navigate to the Lokka project directory
2. Change into the folder `src/mcp/`
3. Run the following command to install the dependencies:

   ```bash
   npm install
   ```

4. After the dependencies are installed, run the following command to build the project:

   ```bash
   npm run build
   ```

5. When the build is complete, you will see a `main.js` file and other compiled files in the `src/mcp/build/` folder

## Configuring the Agent

### Claude Desktop

1. In Claude Desktop, open the settings by clicking on the hamburger icon in the top left corner
2. Select **File > Settings** (or press `Ctrl + ,`)
3. In the **Developer** tab, click **Edit Config**
4. This opens explorer, edit `claude_desktop_config.json` in your favorite text editor
5. Add the following configuration to the file, using the information from the Overview blade of the Entra application you created earlier:

**Note:** On Windows the path needs to be escaped with `\\` or use `/` instead of `\`.

Example paths:

- Windows: `C:\\Users\\<username>\\Documents\\lokka\\src\\mcp\\build\\main.js`
- Windows (alternative): `C:/Users/<username>/Documents/lokka/src/mcp/build/main.js`
- macOS/Linux: `/Users/<username>/Documents/lokka/src/mcp/build/main.js`

**Tip:** Right-click on `build/main.js` in VS Code and select "Copy path" to copy the full path.

```json
{
  "mcpServers": {
      "Lokka-Microsoft": {
          "command": "node",
          "args": [
              "<absolute-path-to-main.js>/src/mcp/build/main.js"
          ],
          "env": {
            "TENANT_ID": "<tenant-id>",
            "CLIENT_ID": "<client-id>",
            "CLIENT_SECRET": "<client-secret>"
          }
      }
  }
}
```

6. Exit Claude Desktop and restart it

**Important:** Every time you make changes to the code or configuration, you need to restart Claude Desktop for the changes to take effect.

**Note:** In Windows, Claude doesn't exit when you close the window—it runs in the background. You can find it in the system tray. Right-click on the icon and select "Quit" to exit the application completely.

### VS Code

For VS Code configuration, refer to the [installation documentation](../website/docs/install.mdx) for specific setup instructions.

## Testing the Agent

### Testing with Claude Desktop

1. Open the Claude Desktop application
2. In the chat window on the bottom right, you should see a hammer icon if the configuration is correct
3. Now you can start querying your Microsoft tenant using the Lokka agent tool

**Sample queries you can try:**

- "Get all the users in my tenant"
- "Show me the details for John Doe"
- "Change John's department to IT" *(Requires `User.ReadWrite.All` permission to be granted)*

### Testing with MCP Inspector

MCP Inspector is a tool that allows you to test and debug your MCP server directly (without an LLM). It provides a user interface to send requests to the server and view the responses.

See the [MCP Inspector documentation](https://modelcontextprotocol.io/docs/tools/inspector) for more information.

To use MCP Inspector with Lokka:

```bash
npx @modelcontextprotocol/inspector node path/to/server/main.js args...
```

Example:

```bash
npx @modelcontextprotocol/inspector node src/mcp/build/main.js
```

## Development Workflow

1. Make your changes to the TypeScript source files in `src/mcp/src/`
2. Build the project: `npm run build`
3. Restart Claude Desktop or your MCP client
4. Test your changes using either Claude Desktop or MCP Inspector

## Contributing

When contributing to the project:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly using both Claude Desktop and MCP Inspector
5. Submit a pull request with a clear description of your changes

For more information about the project architecture and implementation details, see the source code documentation in the `src/mcp/src/` directory.

# DEMO: Using SPHttpClient to talk to SharePoint

In this demo you will create a SharePoint Framework (SPFx) web part that will get and display data from a SharePoint list.

1. Install the project dependencies:
    1. Open a command prompt and navigate to the folder that contains this demo.
    1. Execute the following command:

        ```shell
        npm install
        ```

1. Start the local web server and test the web part in the hosted workbench:

    ```shell
    gulp serve
    ```

1. The browser will loads the local workbench, but you can not use this for testing because there is no SharePoint context in the local workbench. Instead, navigate to the SharePoint Online site where you created the **Countries** list, and load the hosted workbench at **https://[sharepoint-online-site]/_layouts/workbench.aspx**.

1. Add the web part to the page: Select the **Add a new web part** control...

    ![Screenshot of the SharePoint workbench](../../Images/add-webpart-01.png)

    ...then select the expand toolbox icon in the top-right...

    ![Screenshot of the SharePoint workbench](../../Images/add-webpart-02.png)

    ...and select the **SPFxHttpClientDemo** web part to add the web part to the page:

    ![Screenshot of the SharePoint workbench toolbox](../../Images/add-webpart-03.png)

1. The web part will appear on the page with a single button and no data in the list:

    ![Screenshot of the web part with no data](../../Images/add-webpart-04.png)

1. Select the **Get Countries** button and notice the list will display the data from the SharePoint REST API:

    ![Screenshot of the web part with data](../../Images/get-items-sp.png)

1. Stop the local web server by pressing <kbd>CTRL</kbd>+<kbd>C</kbd> in the console/terminal window.

## Suggested files to Explore in "how it works"

- **./src/webparts/spFxContent/SpFxContentWebPart.ts**
- **./src/webparts/spFxContent/components/ISpFxContentProps.ts**
- **./src/webparts/spFxContent/components/SpFxContent.tsx**
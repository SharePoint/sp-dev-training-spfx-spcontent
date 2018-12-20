# DEMO: Using Mocks to Simulate SharePoint Data

In this exercise, you will extend the SPFx project from the previous exercise to add logic so mock data is used when the web part is run from the local workbench.

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

1. The browser will loads the local workbench, but you can not use this for complete testing because there is no SharePoint context in the local workbench. Instead, navigate to the SharePoint Online site where you created the **Countries** list, and load the hosted workbench at **https://[sharepoint-online-site]/_layouts/workbench.aspx**.

1. Add the web part to the page: Select the **Add a new web part** control...

    ![Screenshot of the SharePoint workbench](../../Images/add-webpart-01.png)

    ...then select the expand toolbox icon in the top-right...

    ![Screenshot of the SharePoint workbench](../../Images/add-webpart-02.png)

    ...and select the **SPFxHttpClientContent** web part to add the web part to the page:

    ![Screenshot of the SharePoint workbench toolbox](../../Images/add-webpart-03.png)

1. The web part will appear on the page with a single button and no data in the list:

    ![Screenshot of the web part with all buttons](../../Images/all-buttons.png)

1. Select the **Get Countries** button and examine the results returned.
1. Go back to the browser with the local workbench loaded. Notice that when you select the **Get Countries** button, you see the mock data returned:

    ![Screenshot of mock data in the web part](../../Images/local-workbench-02.png)

1. Also notice that if you go back to the hosted workbench in a SharePoint Online site, the web part works as it did prior to adding mock data as it uses the live SharePoint REST API in a real SharePoint environment.

## Suggested files to Explore in "how it works"

- **./src/webparts/spFxContent/SpFxContentWebPart.ts**
- **./src/webparts/spFxContent/components/ISpFxContentProps.ts**
- **./src/webparts/spFxContent/components/SpFxContent.tsx**
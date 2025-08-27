module.exports = async function (context, req) {
    context.log('Contact function processed a request.');

    const responseMessage = req.query.name || req.body && req.body.name?
        "Hello, " + (req.query.name || req.body.name) + ". This HTTP triggered function executed successfully."

        : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };
};
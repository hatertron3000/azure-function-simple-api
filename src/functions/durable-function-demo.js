const { EmailClient, KnownEmailSendStatus } = require('@azure/communication-email')
const { app, output } = require('@azure/functions')
const { randomUUID } = require('crypto')
const df = require('durable-functions')
const crypto = require("crypto")

const emailConnectionString = process.env.COMMUNICATION_SERVICES_CONNECTION_STRING
const senderAddress = process.env.SENDER_ADDRESS
const recipientAddress = process.env.RECIPIENT_ADDRESS

const sqlOutput = output.sql({
    commandText: 'dbo.ToDo',
    connectionStringSetting: 'SqlConnectionString'
})

const emailClient = new EmailClient(emailConnectionString)

df.app.orchestration('durable-function-demoOrchestrator', function* (context) {
    const outputs = []
    const input = {
        context,
        now: Date.now()
    }


    const record = outputs.push(yield context.df.callActivity('sql-output', Date.now()))
    outputs.push(yield context.df.callActivity('send-email', record))
    return outputs
})

df.app.activity('sql-output', {
    extraOutputs: [sqlOutput],
    handler: (now, context) => {
        try {
            const record = {
                Id: randomUUID(),
                title: now,
                completed: false,
                url: ""
            }
            context.extraOutputs.set(sqlOutput, record)
            return record
        } catch (err) {
            console.error(err)
            return err
        }
    },
})

df.app.activity('send-email', {
    handler: async () => {
        try {
            const emailMessage = {
                senderAddress,
                content: {
                    subject: "Email from your simple API on Azure",
                    plainText: "This email message is sent from Azure Communication Services Email using the JavaScript SDK."
                },
                recipients: {
                    to: [
                        {
                            address: recipientAddress,
                            displayName: "John Doe"
                        }
                    ]
                }
            }

            const poller = await emailClient.beginSend(emailMessage)

            if (!poller.getOperationState().isStarted) {
                throw "Poller was not started"
            }
            
            let timeElapsed = 0
            while (!poller.isDone()) {
                poller.poll()

                await new Promise(resolve => setTimeout(resolve, POLLER_WAIT_TIME * 1000))
                timeElapsed += POLLER_WAIT_TIME

                if (timeElapsed > 60 * POLLER_WAIT_TIME) {
                    throw "Polling timed out."
                }
            }

            if (poller.getResult().status === KnownEmailSendStatus.Succeeded) {
                context.log(`Email sent in ${timeElapsed} seconds at ${timer}ms`)
                clearInterval(timerInterval)

                return poller.getResult().status

            } else {
                const err = poller.getResult().error

                throw err
            }
        } catch (err) {
            
            return err
        }
    },
})

app.http('durable-function-demoHttpStart', {
    route: 'orchestrators/durable-function-demoOrchestrator',
    extraInputs: [df.input.durableClient()],
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        const client = df.getClient(context)
        const body = await request.text()
        const instanceId = await client.startNew('durable-function-demoOrchestrator', { input: body })

        context.log(`Started orchestration with ID = '${instanceId}'.`)

        return client.createCheckStatusResponse(request, instanceId)
    },
})
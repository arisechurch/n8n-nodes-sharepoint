import { INodeType, INodeTypeDescription } from "n8n-workflow"
import { listOptions, run, siteOptions } from "./GenericFunctions"
import { execute } from "../common"

export class MicrosoftSharepoint implements INodeType {
  description: INodeTypeDescription = {
    displayName: "Microsoft Sharepoint",
    name: "microsoftSharepoint",
    icon: "file:sharepoint.svg",
    group: ["transform"],
    version: 1,
    description: "Microsoft Sharepoint API",
    defaults: {
      name: "Microsoft Sharepoint",
    },
    inputs: ["main"],
    outputs: ["main"],
    credentials: [
      {
        name: "microsoftSharepointApi",
        required: true,
      },
    ],
    properties: [
      {
        displayName: "Sharepoint Site Name or ID",
        name: "site",
        type: "options",
        description:
          'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>',
        required: true,
        typeOptions: {
          loadOptionsMethod: "getSites",
        },
        default: "",
      },
      {
        displayName: "Resource",
        name: "resource",
        type: "options",
        noDataExpression: true,
        default: "lists",
        options: [
          {
            name: "List",
            value: "lists",
          },
        ],
      },
      {
        displayName: "List Name or ID",
        name: "list",
        type: "options",
        description:
          'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>',
        displayOptions: {
          hide: {
            site: [""],
          },
          show: {
            resource: ["lists"],
          },
        },
        typeOptions: {
          loadOptionsMethod: "getLists",
        },
        default: "",
      },
      {
        displayName: "Method",
        name: "method",
        type: "options",
        noDataExpression: true,
        default: "GET",
        options: [
          {
            name: "GET",
            value: "GET",
          },
          {
            name: "POST",
            value: "POST",
          },
          {
            name: "PATCH",
            value: "PATCH",
          },
          {
            name: "PUT",
            value: "PUT",
          },
          {
            name: "DELETE",
            value: "DELETE",
          },
        ],
      },
      {
        displayName: "Path",
        name: "path",
        description: "URL path",
        type: "string",
        required: true,
        default: "/",
      },
      {
        displayName: "Body",
        name: "body",
        description: "JSON body to send",
        type: "json",
        displayOptions: {
          hide: {
            method: ["GET", "DELETE"],
          },
        },
        default: "",
      },
    ],
  }

  methods = {
    loadOptions: {
      getSites: execute(siteOptions),
      getLists: execute(listOptions),
    },
  }

  execute = execute(run)
}

import type { ICredentialType, INodeProperties } from "n8n-workflow"

export class MicrosoftSharepointApi implements ICredentialType {
  name = "microsoftSharepointApi"

  extends = ["microsoftOAuth2Api"]

  displayName = "Microsoft Sharepoint OAuth2 API"

  properties: INodeProperties[] = [
    {
      displayName: "Scope",
      name: "scope",
      type: "hidden",
      default: "openid offline_access Sites.Manage.All",
    },
  ]
}

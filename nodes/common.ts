import { Tag } from "@effect/data/Context"
import { TaggedClass } from "@effect/data/Data"
import { identity } from "@effect/data/Function"
import type { Option } from "@effect/data/Option"
import type { Effect } from "@effect/io/Effect"
import type {
  IDataObject,
  IExecuteFunctions,
  IExecuteSingleFunctions,
  ILoadOptionsFunctions,
  INodeExecutionData,
  JsonObject,
} from "n8n-workflow"
import { NodeApiError } from "n8n-workflow"
import type { OptionsWithUri } from "request"

export const N8NId = Symbol.for("N8N")
export interface N8N {
  readonly _id: typeof N8NId
}
export const N8N = Tag<
  N8N,
  IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions
>()

export function execute<E, A>(
  effect: Effect<N8N, E, A>,
): (
  this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions,
) => Promise<A> {
  return function () {
    return effect.provideService(N8N, this).runPromise
  }
}

export function apiRequest<A = any>(
  method: string,
  resource: string,
  {
    qs = {},
    body,
    uri,
    headers = {},
  }: {
    body?: any
    qs?: IDataObject
    uri?: string
    headers?: IDataObject
  } = {},
): Effect<N8N, NodeApiError, A> {
  const options: OptionsWithUri = {
    headers: {
      ...headers,
      "Content-Type": "application/json",
    },
    method,
    body,
    qs,
    uri: uri || `https://graph.microsoft.com/v1.0${resource}`,
    json: true,
  }

  return Do($ => {
    const n8n = $(N8N)

    return $(
      Effect.tryCatchPromise(
        () =>
          n8n.helpers.requestOAuth2.call(
            n8n,
            "microsoftSharepointApi",
            options,
          ),
        error =>
          new NodeApiError(n8n.getNode(), error as JsonObject, error as any),
      ),
    )
  })
}

export interface Response<A> {
  readonly value: A
}

export class NoSuchParam extends TaggedClass("NoSuchParam")<{
  readonly name: string
}> {
  readonly message = `param not set: ${this.name}`
}

export const getOptionalParam = <A = string>(
  name: string,
  i = 0,
): Effect<N8N, never, Option<A>> =>
  Do($ => {
    const n8n = $(N8N)
    const getParam = Option.liftThrowable(n8n.getNodeParameter.bind(n8n))
    return (getParam(name, i) as Option<A>).filter(_ => !!_)
  })

export const getParam = <A = string>(name: string, i = 0) =>
  getOptionalParam<A>(name, i)
    .flatMap(identity)
    .mapError(_ => new NoSuchParam({ name }))

export function wrapJson(json: any): INodeExecutionData[] {
  json = json.value ?? json

  if (Array.isArray(json)) {
    return json.map(json => ({ json }))
  }

  return [{ json }]
}

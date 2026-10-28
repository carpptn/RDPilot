using System.Drawing;
using System.Reflection;

namespace RDPilot.SelfTests.Behavior;
internal static partial class BehavioralTests
{
    static void TestContinuousActionSchema()
    {
        var continuous = RDPilot.OpenAI.PromptAndRequestFactory.ControlActionTypes("continuous");
        var finite = RDPilot.OpenAI.PromptAndRequestFactory.ControlActionTypes("finite");
        Assert(!continuous.Contains("done", StringComparer.OrdinalIgnoreCase), "continuous action schema still exposes done");
        Assert(finite.Contains("done", StringComparer.OrdinalIgnoreCase), "finite action schema lost done");
    }

    static void TestSafeCandidateBatchBarriers()
    {
        List<ActionDto> launchChain = [new ActionDto
        {
            Type = "keys",
            Keys = ["win+r"]
        }, new ActionDto
        {
            Type = "type_text",
            Text = "mspaint"
        }, new ActionDto
        {
            Type = "keys",
            Keys = ["enter"]
        }, new ActionDto
        {
            Type = "wait",
            WaitSeconds = 2
        }, new ActionDto
        {
            Type = "keys",
            Keys = ["escape"]
        }

        ];
        var safe = RDPilot.Control.BatchPlanner.SafeBatchFollowUps(launchChain);
        Assert(safe.Select(action => action.Type).SequenceEqual(["type_text", "keys", "wait"]), "safe launch chain did not stop at its wait barrier");
        var afterCommit = RDPilot.Control.BatchPlanner.SafeBatchFollowUps([new ActionDto { Type = "keys", Keys = ["enter"] }, new ActionDto { Type = "keys", Keys = ["escape"] }]);
        Assert(afterCommit.Length == 0, "state-dependent key was queued after a committing key");
        List<ActionDto> formChain = [new ActionDto
        {
            Type = "keys",
            Keys = ["ctrl", "a"]
        }, new ActionDto
        {
            Type = "type_text",
            Text = "1100"
        }, new ActionDto
        {
            Type = "keys",
            Keys = ["tab"]
        }, new ActionDto
        {
            Type = "type_text",
            Text = "540"
        }, new ActionDto
        {
            Type = "keys",
            Keys = ["enter"]
        }

        ];
        var safeForm = RDPilot.Control.BatchPlanner.SafeBatchFollowUps(formChain);
        Assert(safeForm.Select(action => action.Type).SequenceEqual(["type_text", "keys", "type_text", "keys"]), "deterministic focused form-edit sequence was split by an observation barrier");
        List<ActionDto> observedTurnClicks = [new ActionDto
        {
            Type = "click",
            XPx = 100,
            YPx = 100,
            Note = "Move up through the corridor and then right toward the goal.",
            PlannedInputs = ["ArrowUp", "ArrowRight", "ArrowRight"],
            PlanConfidence = 0.8
        }, new ActionDto
        {
            Type = "click",
            XPx = 120,
            YPx = 100,
            Note = "Continue toward the upper terminal."
        }, new ActionDto
        {
            Type = "click",
            XPx = 120,
            YPx = 100,
            Note = "Continue toward the upper terminal."
        }

        ];
        var observedTurnFollowUps = RDPilot.Control.BatchPlanner.SafeObservedTurnFollowUps(observedTurnClicks, 12);
        Assert(observedTurnFollowUps.Length == 2 && observedTurnClicks[0].ResolvedTurnInputLabel == "ArrowUp" && observedTurnFollowUps.All(action => action.ResolvedTurnInputLabel == "ArrowRight"), "descriptive notes overrode the positional planned-input mapping");
        var drawOne = new ActionDto
        {
            Type = "drag_path",
            GestureKind = "draw",
            DurationMs = 300,
            Path = [new GesturePointDto
            {
                XPx = 10,
                YPx = 10
            }, new GesturePointDto
            {
                XPx = 20,
                YPx = 20
            }

            ]
        };
        var drawTwo = new ActionDto
        {
            Type = "drag_path",
            GestureKind = "draw",
            DurationMs = 300,
            Path = [new GesturePointDto
            {
                XPx = 30,
                YPx = 30
            }, new GesturePointDto
            {
                XPx = 40,
                YPx = 40
            }

            ]
        };
        Assert(RDPilot.Control.BatchPlanner.SafeBatchFollowUps([drawOne, drawTwo]).Length == 0, "draw strokes were batched without a stable-canvas context");
        var drawBatch = RDPilot.Control.BatchPlanner.SafeBatchFollowUps([drawOne, drawTwo], allowDrawGestureBatch: true);
        Assert(drawBatch.Length == 1 && drawBatch[0] == drawTwo, "independent draw strokes on a stable canvas were not batched");
        var pan = new ActionDto
        {
            Type = "drag_path",
            GestureKind = "pan",
            DurationMs = 300,
            Path = [new GesturePointDto
            {
                XPx = 50,
                YPx = 50
            }, new GesturePointDto
            {
                XPx = 60,
                YPx = 60
            }

            ]
        };
        Assert(RDPilot.Control.BatchPlanner.SafeBatchFollowUps([drawOne, pan], allowDrawGestureBatch: true).Length == 0, "state-changing pan gesture was accepted into a draw batch");
        Assert(RDPilot.Control.ControlBatchService.AllowsStableCanvasDrawBatch(new UiPromptContext("Paint", "mspaint", "canvas", null, null, null), "static_ui"), "recognized drawing application did not enable stable-canvas draw batching");
        Assert(!RDPilot.Control.ControlBatchService.AllowsStableCanvasDrawBatch(new UiPromptContext("Game", "game", "canvas", null, null, null), "realtime_interaction"), "realtime canvas incorrectly enabled draw batching");
        var unsafeGesture = RDPilot.Control.BatchPlanner.SafeBatchFollowUps([new ActionDto { Type = "type_text", Text = "x" }, new ActionDto { Type = "drag_path", GestureKind = "draw" }]);
        Assert(unsafeGesture.Length == 0, "gesture was queued without a fresh observation");
        var invalidFollowUp = RDPilot.Control.BatchPlanner.SafeBatchFollowUps([new ActionDto { Type = "keys", Keys = ["win"] }, new ActionDto { Type = "type_text" }]);
        Assert(invalidFollowUp.Length == 0, "invalid text input was queued after a valid prelude");
        Assert(RDPilot.Control.ControlBatchService.BatchTransitionDelayMs(launchChain[0], launchChain[1]) >= 200, "text-entry prelude does not give the target time to receive focus");
        Assert(RDPilot.Control.ControlBatchService.BatchTransitionDelayMs(launchChain[1], launchChain[2]) >= 400, "commit key does not give dynamic suggestions time to become ready");
        Assert(RDPilot.Control.ControlBatchService.BatchTransitionDelayMs(launchChain[2], launchChain[3]) == 0, "a terminal wait received a redundant transition delay");
        Assert(RDPilot.Control.ControlBatchService.IsAdaptiveBatchWaitTrigger(launchChain[2]), "a committing key does not enable adaptive terminal waiting");
        Assert(!RDPilot.Control.ControlBatchService.IsAdaptiveBatchWaitTrigger(launchChain[1]), "plain text entry incorrectly enables adaptive terminal waiting");
        Assert(RDPilot.Control.ControlBatchService.IsNavigationBatchPrelude(launchChain[0]) && RDPilot.Control.ControlBatchService.IsNavigationBatchPrelude(new ActionDto { Type = "keys", Keys = ["ctrl+l"] }) && !RDPilot.Control.ControlBatchService.IsNavigationBatchPrelude(new ActionDto { Type = "keys", Keys = ["ctrl+a"] }), "navigation batch preludes are not distinguished from ordinary shortcuts");
        Assert(RDPilot.Control.ControlBatchService.AdaptiveBatchMinimumObservationMs(timeoutMs: 6000, navigationLikeBatch: true) == 6000 && RDPilot.Control.ControlBatchService.AdaptiveBatchMinimumObservationMs(timeoutMs: 6000, navigationLikeBatch: false) == 900, "navigation waits can still settle on a short-lived static loading screen");
        var delayedTurnReactionPolicy = RDPilot.Control.ControlActionPolicy.ResolveTurnNoEffectPolicy(immediateReactionObserved: false);
        Assert(delayedTurnReactionPolicy.ExtendObservation && !delayedTurnReactionPolicy.ReplayInput, "a delayed turn reaction still permits duplicate input injection");
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(1000, 800, 1000, 800);
        var firstTurnClick = new ActionDto
        {
            Type = "click",
            X = 0.2,
            Y = 0.8,
            Note = "move up",
            PlannedInputs = ["ArrowUp", "ArrowUp", "ArrowRight"],
            PlanConfidence = 0.60
        };
        var repeatedTurnClick = new ActionDto
        {
            Type = "click",
            X = 0.2,
            Y = 0.8,
            Note = "move up"
        };
        var mixedTurnKey = new ActionDto
        {
            Type = "keys",
            Keys = ["ArrowRight"]
        };
        Assert(RDPilot.Control.BatchPlanner.SafeBatchFollowUps([firstTurnClick, repeatedTurnClick, mixedTurnKey]).Length == 0, "turn clicks bypassed the explicit observed-batch gate");
        var observedTurnBatch = RDPilot.Control.BatchPlanner.SafeBatchFollowUps([firstTurnClick, repeatedTurnClick, mixedTurnKey], observedTurnBatchLimit: 4);
        Assert(observedTurnBatch.SequenceEqual([repeatedTurnClick, mixedTurnKey]), "a reversible planned click/key route was not accepted with per-input observation");
    }

    static void TestControlActionBatchSchema()
    {
        var schemaJson = System.Text.Json.JsonSerializer.Serialize(RDPilot.OpenAI.PromptAndRequestFactory.ControlActionBatchSchema(1280, 720));
        using var schema = System.Text.Json.JsonDocument.Parse(schemaJson);
        var actions = schema.RootElement.GetProperty("properties").GetProperty("actions");
        Assert(actions.GetProperty("minItems").GetInt32() == 1, "action sequence may be empty");
        Assert(actions.GetProperty("maxItems").GetInt32() >= 1, "action sequence has an invalid bound");
        var variants = actions.GetProperty("items").GetProperty("anyOf");
        Assert(variants.GetArrayLength() > 1, "control actions are not represented by compact variants");
        var keysVariant = variants.EnumerateArray().First(variant => variant.GetProperty("properties").GetProperty("type").GetProperty("enum")[0].GetString() == "keys");
        var keysProperties = keysVariant.GetProperty("properties");
        Assert(keysProperties.TryGetProperty("keys", out _), "keys variant lost its action payload");
        Assert(keysProperties.GetProperty("keys").GetProperty("maxItems").GetInt32() == 32, "directional key actions do not expose the configured extended horizon");
        Assert(!keysProperties.TryGetProperty("x", out _) && !keysProperties.TryGetProperty("text", out _) && !keysProperties.TryGetProperty("path", out _), "keys variant still forces unrelated action fields");
        Assert(keysVariant.GetProperty("required").GetArrayLength() == keysProperties.EnumerateObject().Count(), "strict action variant has optional fields");
        var envelopeProperties = schema.RootElement.GetProperty("properties");
        Assert(envelopeProperties.TryGetProperty("world_state_summary", out _) && envelopeProperties.TryGetProperty("mechanics_hypothesis", out _) && envelopeProperties.TryGetProperty("salient_change_observation", out _) && envelopeProperties.TryGetProperty("short_term_plan", out _) && envelopeProperties.TryGetProperty("plan_status", out _) && envelopeProperties.TryGetProperty("plan_revision_reason", out _) && envelopeProperties.TryGetProperty("planned_inputs", out _) && envelopeProperties.TryGetProperty("plan_waypoint", out _) && envelopeProperties.TryGetProperty("plan_state_id", out _) && envelopeProperties.TryGetProperty("plan_confidence", out _), "ActionBatch envelope lost stateful working-memory fields");
        Assert(envelopeProperties.GetProperty("planned_inputs").GetProperty("maxItems").GetInt32() == 32, "structured turn plans do not preserve the configured extended horizon");
        Assert(schema.RootElement.GetProperty("required").GetArrayLength() == 15, "ActionBatch envelope does not carry shared sequence metadata");
        var rules = RDPilot.OpenAI.PromptAndRequestFactory.BuildSystemRules();
        Assert(rules.Contains("actions[0]", StringComparison.Ordinal), "control prompt does not identify the first executable action");
        Assert(!rules.Contains("Return EXACTLY ONE action per round", StringComparison.Ordinal), "legacy single-action instruction still disables batching");
        Assert(rules.Contains("short_term_plan", StringComparison.Ordinal) && rules.Contains("planned_inputs", StringComparison.Ordinal) && rules.Contains("aggressive batching is the default", StringComparison.Ordinal) && rules.Contains("count the route cell by cell", StringComparison.Ordinal), "control prompt does not connect a conditional plan with bounded execution");
        var temporalRequest = RDPilot.OpenAI.PromptAndRequestFactory.BuildRequestBody("gpt-5.6-luna", rules, "solve the visible board", "", "data:image/jpeg;base64,current-screen", 1280, 720, 0, 0, 0, 0, "data:image/jpeg;base64,current-focus", new Rectangle(100, 100, 600, 400), null, null, new UiPromptContext("Board", "browser", null, null, null, null), false, null, false, "finite", "data:image/jpeg;base64,previous-focus", "data:image/jpeg;base64,reference-focus", [new TurnChangeImagePair("data:image/jpeg;base64,change-before", "data:image/jpeg;base64,change-after", null, null, 1)], 10000);
        var temporalRequestJson = System.Text.Json.JsonSerializer.Serialize(temporalRequest);
        Assert(temporalRequestJson.Contains("PREVIOUS_TURN_STATE_IMAGE", StringComparison.Ordinal) && temporalRequestJson.Contains("TURN_REFERENCE_IMAGE", StringComparison.Ordinal) && temporalRequestJson.Contains("current visual epoch only", StringComparison.Ordinal) && temporalRequestJson.Contains("CURRENT_FOCUS_IMAGE", StringComparison.Ordinal) && temporalRequestJson.Contains("SALIENT_CHANGE_REGION_1_BEFORE", StringComparison.Ordinal) && temporalRequestJson.Contains("SALIENT_CHANGE_REGION_1_AFTER", StringComparison.Ordinal), "control request did not label temporal turn images");
        Assert(temporalRequestJson.Contains("\"max_output_tokens\":10000", StringComparison.Ordinal), "salient control request lost its dynamic output budget");
        var envelope = System.Text.Json.JsonSerializer.Serialize(new ActionBatchDto { Actions = [new ActionDto { Type = "keys", Keys = ["win"] }, new ActionDto { Type = "type_text", Text = "Paint" }], PlannedInputs = ["ArrowUp", "ArrowLeft"], PlanWaypoint = "visible opening", PlanStateId = "S3", PlanConfidence = 0.9 });
        var responseJson = System.Text.Json.JsonSerializer.Serialize(new { output = new[] { new { content = new[] { new { type = "output_text", text = envelope } } } } });
        using var response = System.Text.Json.JsonDocument.Parse(responseJson);
        Assert(RDPilot.OpenAI.OpenAiResponsesService.TryParseControlActionSequence(response.RootElement, out var parsedActions, out var payloadCount, out var legacyPayload), "explicit action sequence could not be parsed from a Responses payload");
        Assert(parsedActions.Count == 2 && payloadCount == 1 && !legacyPayload && parsedActions[0].PlannedInputs?.SequenceEqual(["ArrowUp", "ArrowLeft"]) == true && parsedActions[0].PlanStateId == "S3" && parsedActions[0].PlanConfidence == 0.9, "explicit action sequence lost actions or was treated as legacy output");
    }

    static void TestControlActionStreaming()
    {
        var actionJson = "{\"actions\":[{\"type\":\"keys\",\"keys\":[\"win\"]}],\"confidence\":0.99,\"note\":\"open Start\",\"recovery_strategy_id\":null,\"recovery_strategy_step\":null}";
        var eventJson = System.Text.Json.JsonSerializer.Serialize(new { type = "response.output_text.done", text = actionJson });
        string? responseId = "resp_test";
        Assert(RDPilot.OpenAI.OpenAiResponsesService.TryHandleControlStreamEvent(eventJson, allowEarlyAccept: true, ref responseId, out var raw, out var actions, out var completed), "complete streamed ActionBatch was not accepted");
        Assert(actions is { Count: 1 } && actions[0].Type == "keys" && actions[0].Confidence == 0.99 && actions[0].Note == "open Start", "streamed ActionBatch lost its first action");
        Assert(!completed && raw.Contains("rdpilot_stream_early_accept", StringComparison.Ordinal), "early stream acceptance did not preserve a replayable response");
        Assert(!RDPilot.OpenAI.OpenAiResponsesService.TryHandleControlStreamEvent(eventJson, allowEarlyAccept: false, ref responseId, out _, out _, out _), "stream was accepted early while previous_response_id state required completion");
    }

    static void TestWebSearchIntegration()
    {
        var root = typeof(RDPilotApplication);
        var allowField = TestState.Field(root, "AllowWebSearch", BindingFlags.Static | BindingFlags.NonPublic)!;
        var countField = TestState.Field(root, "RunWebSearchCalls", BindingFlags.Static | BindingFlags.NonPublic)!;
        var idsField = TestState.Field(root, "ReportedWebSearchCallIds", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originalAllow = allowField.GetValue(null);
        var originalCount = countField.GetValue(null);
        var ids = (HashSet<string>)idsField.GetValue(null)!;
        try
        {
            Assert(originalAllow is false, "web search is not disabled by default");
            allowField.SetValue(null, true);
            countField.SetValue(null, 0);
            ids.Clear();
            var request = new Dictionary<string, object>();
            RDPilot.OpenAI.PromptAndRequestFactory.AddWebSearchOptions(request);
            var requestJson = System.Text.Json.JsonSerializer.Serialize(request);
            Assert(requestJson.Contains("\"type\":\"web_search\"", StringComparison.Ordinal) && requestJson.Contains("\"tool_choice\":\"auto\"", StringComparison.Ordinal), "enabled web search was not added to the request");
            var rules = RDPilot.OpenAI.PromptAndRequestFactory.BuildSystemRules();
            Assert(rules.Contains("Web search is available", StringComparison.Ordinal) && rules.Contains("Do not search for ordinary UI operation", StringComparison.Ordinal), "enabled web search was not constrained in the control prompt");
            var searchingEvent = System.Text.Json.JsonSerializer.Serialize(new { type = "response.web_search_call.searching", item_id = "ws_test" });
            string? responseId = "resp_web";
            Assert(!RDPilot.OpenAI.OpenAiResponsesService.TryHandleControlStreamEvent(searchingEvent, allowEarlyAccept: false, ref responseId, out _, out _, out _), "web search progress event unexpectedly completed the response");
            var completedResponse = System.Text.Json.JsonSerializer.Serialize(new { output = new object[] { new { type = "web_search_call", id = "ws_test", action = new { type = "search", query = "current test fact" } } } });
            using var completedDocument = System.Text.Json.JsonDocument.Parse(completedResponse);
            RDPilot.OpenAI.OpenAiResponsesService.RecordWebSearchMetrics(completedDocument.RootElement);
            Assert((int)countField.GetValue(null)! == 1, "one streamed web search call was counted more than once");
            allowField.SetValue(null, false);
            var disabledRequest = new Dictionary<string, object>();
            RDPilot.OpenAI.PromptAndRequestFactory.AddWebSearchOptions(disabledRequest);
            Assert(disabledRequest.Count == 0, "disabled web search still modified the request");
        }
        finally
        {
            allowField.SetValue(null, originalAllow);
            countField.SetValue(null, originalCount);
            ids.Clear();
        }
    }

    static void TestTemperatureSupport()
    {
        foreach (var model in new[]
        {
            "gpt-5.6-luna",
            "gpt-6.1-terra",
            "GPT-6.1-TERRA"
        }

        )
            Assert(!RDPilot.OpenAI.PromptAndRequestFactory.SupportsTemperature(model), $"{model} must omit temperature");
        Assert(RDPilot.OpenAI.PromptAndRequestFactory.SupportsTemperature("gpt-4.1"), "temperature support for older models changed");
    }

    static void TestControlContextRequestOptions()
    {
        var root = typeof(RDPilotApplication);
        var stateField = TestState.Field(root, "UsePreviousResponseState", BindingFlags.Static | BindingFlags.NonPublic)!;
        var contextField = TestState.Field(root, "ControlReasoningContext", BindingFlags.Static | BindingFlags.NonPublic)!;
        var compactionField = TestState.Field(root, "ControlContextCompactionEnabled", BindingFlags.Static | BindingFlags.NonPublic)!;
        var thresholdField = TestState.Field(root, "ControlContextCompactThreshold", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originals = new[]
        {
            stateField.GetValue(null),
            contextField.GetValue(null),
            compactionField.GetValue(null),
            thresholdField.GetValue(null)
        };
        try
        {
            stateField.SetValue(null, true);
            contextField.SetValue(null, "all_turns");
            compactionField.SetValue(null, true);
            thresholdField.SetValue(null, 700_000);
            var control = new Dictionary<string, object>();
            RDPilot.OpenAI.PromptAndRequestFactory.AddReasoningOptions(control, "gpt-5.6-luna", "max", 6000, "control");
            RDPilot.OpenAI.PromptAndRequestFactory.AddControlContextManagement(control, enableContextCompaction: true);
            var controlJson = System.Text.Json.JsonSerializer.Serialize(control);
            Assert(controlJson.Contains("\"context\":\"all_turns\"", StringComparison.Ordinal), "control request did not preserve reasoning across turns");
            Assert(controlJson.Contains("\"context_management\":[{\"type\":\"compaction\",\"compact_threshold\":700000}]", StringComparison.Ordinal), "control request did not use the documented server-side compaction shape");
            var verifier = new Dictionary<string, object>();
            RDPilot.OpenAI.PromptAndRequestFactory.AddReasoningOptions(verifier, "gpt-5.6-luna", "max", 6000, "verify");
            var verifierJson = System.Text.Json.JsonSerializer.Serialize(verifier);
            Assert(verifierJson.Contains("\"context\":\"current_turn\"", StringComparison.Ordinal), "independent verifier inherited the control reasoning context");
            Assert(!verifierJson.Contains("previous_response_id", StringComparison.Ordinal), "independent verifier inherited the control response chain");
        }
        finally
        {
            stateField.SetValue(null, originals[0]);
            contextField.SetValue(null, originals[1]);
            compactionField.SetValue(null, originals[2]);
            thresholdField.SetValue(null, originals[3]);
        }
    }

    static void TestControlContextChainLifecycle()
    {
        var chain = new RDPilot.Control.ControlContextChain("1234567890abcdef", enabled: true, compactionEnabled: true, fallbackLimit: 3);
        Assert(chain.PreviousResponseIdForRequest is null, "new task inherited an old response id");
        chain.RecordResult("resp_first", false, false, false);
        Assert(chain.PreviousResponseIdForRequest == "resp_first" && chain.TurnCount == 1, "completed response did not advance the control chain");
        chain.RecordResult("resp_restarted", true, false, true);
        Assert(chain.PreviousResponseIdForRequest == "resp_restarted" && chain.RestartCount == 1, "checkpoint fallback did not establish a replacement chain");
        chain.RecordResult("resp_no_compaction", false, true, false);
        Assert(chain.Enabled && !chain.CompactionEnabled, "compaction rejection disabled the whole control chain");
        chain.RecordResult("resp_restart_2", true, false, false);
        chain.RecordResult("resp_restart_3", true, false, false);
        Assert(!chain.Enabled && chain.PreviousResponseIdForRequest is null, "repeated response-state failures did not fall back to explicit application history");
    }

    static void TestControlContextFallback()
    {
        var body = new Dictionary<string, object>
        {
            ["model"] = "gpt-5.6-luna",
            ["previous_response_id"] = "resp_old",
            ["context_management"] = new object[]
            {
                new
                {
                    type = "compaction",
                    compact_threshold = 700_000
                }
            },
            ["reasoning"] = new
            {
                effort = "max",
                context = "all_turns"
            },
            ["input"] = new object[]
            {
                new
                {
                    role = "user",
                    content = "checkpoint"
                }
            }
        };
        Assert(RDPilot.OpenAI.OpenAiResponsesService.IsPreviousResponseStateFailure(400, "Invalid previous_response_id: response not found"), "previous-response API error was not classified as a chain failure");
        Assert(RDPilot.OpenAI.OpenAiResponsesService.TryBuildRequestWithoutProperty(body, "previous_response_id", out var fallbackBody), "could not construct a checkpoint retry body");
        var fallbackJson = System.Text.Json.JsonSerializer.Serialize(fallbackBody);
        Assert(!fallbackJson.Contains("previous_response_id", StringComparison.Ordinal) && fallbackJson.Contains("context_management", StringComparison.Ordinal) && fallbackJson.Contains("checkpoint", StringComparison.Ordinal), "checkpoint retry discarded current request state or retained the invalid response id");
        var completed = "{\"id\":\"resp_ok\",\"status\":\"completed\",\"output\":[{\"type\":\"compaction\"}]}";
        var incomplete = "{\"id\":\"resp_bad\",\"status\":\"incomplete\",\"output\":[]}";
        Assert(RDPilot.OpenAI.OpenAiResponsesService.TryGetCompletedResponseId(completed) == "resp_ok", "completed response id was not accepted");
        Assert(RDPilot.OpenAI.OpenAiResponsesService.TryGetCompletedResponseId(incomplete)is null, "incomplete response id was accepted into the chain");
        using var completedDocument = System.Text.Json.JsonDocument.Parse(completed);
        Assert(RDPilot.OpenAI.OpenAiResponsesService.ContainsCompactionItem(completedDocument.RootElement), "server compaction output was not detected");
    }

    static void TestProductiveContinuousCycle()
    {
        var graph = new LoopStateGraph();
        var context = new UiPromptContext("Operations dashboard", "msedge", "status monitor", null, null, null);
        var actions = new Queue<ResolvedActionSnapshot>();
        var states = new[]
        {
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10)
        };
        LoopDetectionAssessment? assessment = null;
        for (var index = 0; index < states.Length; index++)
        {
            ResolvedActionSnapshot? previous = null;
            if (index > 0)
            {
                var type = index % 2 == 0 ? "request_crop" : "wait";
                previous = Snapshot(type, new Point(100 + index, 100));
                previous = previous with
                {
                    SemanticTokens = "monitor status refresh"
                };
                actions.Enqueue(previous);
            }

            assessment = RDPilot.Persistence.RecoveryMemoryService.AssessVisualStateCycle(graph, states[index], states[index], context, index + 1, actions, previous, index == 0 ? double.NaN : 0.1, recordLearning: false, goalMode: "continuous", recurringWorkflowIntent: true);
        }

        Assert(assessment?.IsProductiveCycle == true, "goal-aligned recurring workflow was not classified as productive");
        Assert(assessment!.IsLoop == false, "productive recurring workflow was treated as a harmful loop");
        Assert(assessment.CycleDisposition == "productive", "productive recurring workflow has the wrong disposition");
        var inspectionActions = new Queue<ResolvedActionSnapshot>();
        var openHelp = Snapshot("click", new Point(400, 400))with
        {
            SemanticTokens = "open help instructions"
        };
        var inspectHelp = Snapshot("request_crop", new Point(420, 420))with
        {
            SemanticTokens = "inspect instruction details"
        };
        var returnFromHelp = Snapshot("click", new Point(100, 100))with
        {
            SemanticTokens = "return back to prior workspace"
        };
        inspectionActions.Enqueue(openHelp);
        inspectionActions.Enqueue(inspectHelp);
        inspectionActions.Enqueue(returnFromHelp);
        Assert(RDPilot.Persistence.RecoveryMemoryService.IsCompletedInspectionRoundTrip(inspectionActions, returnFromHelp), "a completed help inspection was not recognized as a productive round-trip");
    }

    static void TestGoalModes()
    {
        Assert(RDPilot.Persistence.RecoveryMemoryService.ClassifyGoalMode("Monitoruj kolejkę nowych zadań") == "continuous", "general monitoring goal was not continuous");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ClassifyGoalMode("Utrzymuj proces aktywny do odwołania") == "continuous", "user-terminated maintenance goal was not continuous");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ClassifyGoalMode("Obserwuj logi przez 10 minut") == "finite", "time-bounded observation goal was not finite");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ClassifyGoalMode("Graj w grę") == "continuous", "open-ended game goal was not continuous");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ClassifyGoalMode("Otwórz zadanie, samodzielnie rozwiąż pierwszy poziom i doprowadź go do stanu ukończenia. Nie przerywaj po częściowym postępie.") == "finite", "explicitly bounded game goal was overridden by open-ended wording");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ClassifyGoalMode("Sprawdzaj okresowo stan kolejki") == "continuous", "general recurring workflow was not continuous");
        Assert(RDPilot.Persistence.RecoveryMemoryService.HasRecurringWorkflowIntent("Sprawdzaj okresowo stan kolejki"), "recurring workflow intent was not recognized");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ClassifyGoalMode("Powtarzaj sprawdzenie 5 razy") == "finite", "count-bounded recurring workflow was not finite");
    }

    static void TestGoalModeOverride()
    {
        Assert(RDPilot.Persistence.RecoveryMemoryService.ResolveGoalMode("Monitoruj kolejkę", "finite") == "finite", "finite override did not win over the heuristic");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ResolveGoalMode("Zamknij Notatnik", "continuous") == "continuous", "continuous override did not win over the heuristic");
        Assert(RDPilot.Persistence.RecoveryMemoryService.ResolveGoalMode("Monitoruj kolejkę", "auto") == "continuous", "auto mode did not delegate to classification");
    }

    static void TestContinuousIdle()
    {
        var wait = new ActionDto
        {
            Type = "wait",
            WaitSeconds = 10
        };
        Assert(RDPilot.Control.ControlActionPolicy.IsExpectedContinuousIdle("continuous", wait, noChange: true), "unchanged continuous wait was treated as stagnation");
        Assert(!RDPilot.Control.ControlActionPolicy.IsExpectedContinuousIdle("finite", wait, noChange: true), "finite wait incorrectly bypassed stagnation");
        Assert(!RDPilot.Control.ControlActionPolicy.IsExpectedContinuousIdle("continuous", new ActionDto { Type = "click", XPx = 10, YPx = 10 }, noChange: true), "ineffective continuous click incorrectly bypassed stagnation");
        Assert(!RDPilot.Control.ControlActionPolicy.IsExpectedContinuousIdle("continuous", wait, noChange: false), "a changed screen was labelled idle");
    }

    static void TestAutoVerificationPolicy()
    {
        var root = typeof(RDPilotApplication);
        var modeField = TestState.Field(root, "VerifyMode", BindingFlags.Static | BindingFlags.NonPublic)!;
        var thresholdField = TestState.Field(root, "SkipVerifyConfidenceThreshold", BindingFlags.Static | BindingFlags.NonPublic)!;
        var earlyField = TestState.Field(root, "VerifyEarlySteps", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originalMode = modeField.GetValue(null);
        var originalThreshold = thresholdField.GetValue(null);
        var originalEarly = earlyField.GetValue(null);
        try
        {
            modeField.SetValue(null, "auto");
            thresholdField.SetValue(null, 0.92);
            earlyField.SetValue(null, 2);
            Assert(!RDPilot.Control.ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 11, new ActionDto { Type = "done", Confidence = 0.99 }), "auto mode did not skip a strong late completion");
            Assert(RDPilot.Control.ControlGoalPolicy.VerificationSkipReason(11, new ActionDto { Type = "done", Confidence = 0.99 }).Contains("high-confidence completion", StringComparison.Ordinal), "verification skip log does not explain the high-confidence threshold");
            Assert(RDPilot.Control.ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 11, new ActionDto { Type = "done", Confidence = 0.7 }), "auto mode skipped a low-confidence completion");
            Assert(RDPilot.Control.ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 1, new ActionDto { Type = "done", Confidence = 0.99 }), "auto mode skipped verification during early steps");
        }
        finally
        {
            modeField.SetValue(null, originalMode);
            thresholdField.SetValue(null, originalThreshold);
            earlyField.SetValue(null, originalEarly);
        }
    }

    static void TestOutputRetriesFollowReasoningFallback()
    {
        object ladderBody = new Dictionary<string, object?>
        {
            ["model"] = "gpt-5.6-luna",
            ["max_output_tokens"] = 6000,
            ["previous_response_id"] = "resp_previous",
            ["reasoning"] = new Dictionary<string, object?>
            {
                ["effort"] = "max",
                ["context"] = "all_turns"
            }
        };
        var expectedEfforts = new[]
        {
            "max",
            "xhigh",
            "high",
            "medium",
            "low"
        };
        for (var index = 1; index < expectedEfforts.Length; index++)
        {
            var canRetry = RDPilot.OpenAI.OpenAiResponsesService.TryBuildMaxOutputRetryBody(ladderBody, out var nextBody, out var nextMaxTokens, out var nextEffort, out _);
            Assert(canRetry, $"retry stopped before reaching {expectedEfforts[index]}");
            Assert(nextMaxTokens == 6000, $"retry enlarged the token budget instead of lowering effort at {expectedEfforts[index]}");
            Assert(nextEffort == expectedEfforts[index], $"retry did not step down to {expectedEfforts[index]}");
            var nextReasoning = (System.Text.Json.Nodes.JsonObject)((System.Text.Json.Nodes.JsonObject)nextBody)["reasoning"]!;
            Assert(nextReasoning["effort"]?.GetValue<string>() == expectedEfforts[index], $"retry body did not contain {expectedEfforts[index]}");
            Assert(nextReasoning["context"]?.GetValue<string>() == "all_turns", "effort fallback lost persisted reasoning context");
            Assert(((System.Text.Json.Nodes.JsonObject)nextBody)["previous_response_id"]?.GetValue<string>() == "resp_previous", "effort fallback lost the previous response id");
            ladderBody = nextBody;
        }

        var exhausted = RDPilot.OpenAI.OpenAiResponsesService.TryBuildMaxOutputRetryBody(ladderBody, out _, out var exhaustedMaxTokens, out var exhaustedEffort, out _);
        Assert(!exhausted, "retry continued below the low effort floor");
        Assert(exhaustedMaxTokens == 6000, "exhausted retry changed the final token cap");
        Assert(exhaustedEffort == "low", "exhausted retry lost the low effort marker");
        var bodyWithoutEffort = new Dictionary<string, object?>
        {
            ["model"] = "model-without-effort",
            ["max_output_tokens"] = 1000
        };
        Assert(RDPilot.OpenAI.OpenAiResponsesService.TryBuildMaxOutputRetryBody(bodyWithoutEffort, out _, out var expandedTokens, out _, out _) && expandedTokens > 1000, "a request without a reasoning ladder cannot expand its output budget");
    }
}

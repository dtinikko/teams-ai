﻿
namespace Microsoft.Teams.AI.State
{
    /// <summary>
    /// Temporary state.
    /// </summary>
    /// <remarks>
    /// Inherit a new class from this base abstract class to strongly type the applications temp state.
    /// </remarks>
    public class TempState : Record
    {
        /// <summary>
        /// Name of the input property.
        /// </summary>
        public const string InputKey = "input";

        /// <summary>
        /// Name of the output property.
        /// </summary>
        public const string OutputKey = "output";

        /// <summary>
        /// Name of the history property.
        /// </summary>
        public const string HistoryKey = "history";

        /// <summary>
        /// Name of the action outputs property.
        /// </summary>
        public const string ActionOutputsKey = "actionOutputs";

        /// <summary>
        /// Name of the auth tokens property.
        /// </summary>
        public const string AuthTokenKey = "authTokens";


        /// <summary>
        /// Name of the duplicate token exchange property
        /// </summary>
        public const string DuplicateTokenExchangeKey = "duplicateTokenExchange";

        /// <summary>
        /// Creates a new instance of the <see cref="TempState"/> class.
        /// </summary>
        public TempState() : base()
        {
            this[InputKey] = string.Empty;
            this[OutputKey] = string.Empty;
            this[HistoryKey] = string.Empty;
            this[ActionOutputsKey] = new Dictionary<string, string>();
            this[AuthTokenKey] = new Dictionary<string, string>();
            this[DuplicateTokenExchangeKey] = false;
        }

        /// <summary>
        /// Input passed to an AI prompt
        /// </summary>
        public string Input
        {
            get => Get<string>(InputKey)!;
            set => Set(InputKey, value);
        }

        // TODO: This is currently not used, should store AI prompt/function output here
        /// <summary>
        /// Output returned from an AI prompt or function
        /// </summary>
        public string Output
        {
            get => Get<string>(OutputKey)!;
            set => Set(OutputKey, value);
        }


        /// <summary>
        /// Formatted conversation history for embedding in an AI prompt
        /// </summary>
        public string History
        {
            get => Get<string>(HistoryKey)!;
            set => Set(HistoryKey, value);
        }

        /// <summary>
        /// All outputs returned from the action sequence that was executed.
        /// </summary>
        public Dictionary<string, string> ActionOutputs
        {
            get => Get<Dictionary<string, string>>(ActionOutputsKey)!;
            set => Set(ActionOutputsKey, value);
        }

        /// <summary>
        /// All tokens acquired after sign-in for current activity
        /// </summary>
        public Dictionary<string, string> AuthTokens
        {
            get => Get<Dictionary<string, string>>(AuthTokenKey)!;
            set => Set(AuthTokenKey, value);
        }

        /// <summary>
        /// Whether current token exchange is a duplicate one
        /// </summary>
        public bool DuplicateTokenExchange
        {
            get => Get<bool>(DuplicateTokenExchangeKey)!;
            set => Set(DuplicateTokenExchangeKey, value);
        }
    }
}

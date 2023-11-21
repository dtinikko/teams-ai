
namespace Microsoft.Teams.AI.Exceptions
{
    /// <summary>
    /// Exception thrown when user authentication fails.
    /// </summary>
    public class AuthException : Exception
    {
        /// <summary>
        /// The reason for the exception.
        /// </summary>
        public string Reason { get; }

        /// <summary>
        /// Construct a new AuthException.
        /// </summary>
        /// <param name="message">The exception message.</param>
        /// <param name="reason">The reason for the exception.</param>
        public AuthException(string message, string reason = AuthExceptionReason.Other) : base(message)
        {
            this.Reason = reason;
        }
    }

    /// <summary>
    /// A list of possible reasons for an AuthException.
    /// </summary>
    public class AuthExceptionReason
    {
        /// <summary>
        /// Cannot initiate authentication flow with incomming activity.
        /// </summary>
        public const string InvalidActivity = "invalidActivity";

        /// <summary>
        /// Authentication flow completed without a token.
        /// </summary>
        public const string CompletionWithoutToken = "completionWithoutToken";

        /// <summary>
        /// Other error.
        /// </summary>
        public const string Other = "other";

    }
}

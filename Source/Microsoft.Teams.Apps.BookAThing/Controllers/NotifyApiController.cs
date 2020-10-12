// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.ApplicationInsights;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Teams.Apps.BookAThing.Cards;
using Microsoft.Teams.Apps.BookAThing.Common;
using Microsoft.Teams.Apps.BookAThing.Common.Models;
using Microsoft.Teams.Apps.BookAThing.Common.Models.Request;
using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;
using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
using Microsoft.Teams.Apps.BookAThing.Common.Providers;
using Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage;
using Microsoft.Teams.Apps.BookAThing.Helpers;
using Microsoft.Teams.Apps.BookAThing.Models;
using Microsoft.Teams.Apps.BookAThing.Providers.Storage;
using Microsoft.Teams.Apps.BookAThing.Resources;
using Newtonsoft.Json;

namespace Microsoft.Teams.Apps.BookAThing.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class NotifyApiController : ControllerBase
    {
        private readonly IBotFrameworkHttpAdapter _adapter;
        private readonly string _appId;
        private readonly ConcurrentDictionary<string, ConversationReference> _conversationReferences;

        private readonly TelemetryClient telemetryClient;
        private readonly IActivityStorageProvider activityStorageProvider;

        /// <summary>
        /// Storage provider to perform insert, update and delete operation on UserFavorites table.
        /// </summary>
        private readonly IFavoriteStorageProvider favoriteStorageProvider;

        /// <summary>
        /// Provider for exposing methods required to perform meeting creation.
        /// </summary>
        private readonly IMeetingProvider meetingProvider;

        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        private readonly IMeetingHelper meetingHelper;

        private readonly ITokenHelper tokenHelper;



        public NotifyApiController(IBotFrameworkHttpAdapter adapter, IConfiguration configuration, ConcurrentDictionary<string, ConversationReference> conversationReferences, TelemetryClient telemetryClient, IActivityStorageProvider activityStorageProvider, IUserConfigurationStorageProvider userConfigurationStorageProvider, IFavoriteStorageProvider favoriteStorageProvider, IMeetingProvider meetingProvider, IMeetingHelper meetingHelper, ITokenHelper tokenHelper)
        {
            _adapter = adapter;
            _conversationReferences = conversationReferences;
            _appId = configuration["MicrosoftAppId"];

            this.telemetryClient = telemetryClient;
            this.activityStorageProvider = activityStorageProvider;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
            this.favoriteStorageProvider = favoriteStorageProvider;
            this.meetingProvider = meetingProvider;
            this.meetingHelper = meetingHelper;
            this.tokenHelper = tokenHelper;
        }

        [HttpPost]
        public async Task<IActionResult> SubmitTaskForIOS([FromBody] MeetingViewModel valuesFromTaskModule)
        {
            var conversationReference = _conversationReferences[valuesFromTaskModule.UserAdObjectId];
            await ((BotAdapter)_adapter).ContinueConversationAsync(_appId,conversationReference, async (turnContext, cancellationToken) =>
            {
                var activity = turnContext.Activity;

                var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
                if (userConfiguration == null)
                {
                    this.telemetryClient.TrackTrace("User configuration is null in task module submit action.");
                    await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                    return;
                }

                var message = valuesFromTaskModule.Text;
                var replyToId = valuesFromTaskModule.ReplyTo;

                if (message.Equals(BotCommands.MeetingFromTaskModule, StringComparison.OrdinalIgnoreCase))
                {
                    var attachment = SuccessCard.GetSuccessAttachment(valuesFromTaskModule, userConfiguration.WindowsTimezone);
                    var activityFromStorage = await this.activityStorageProvider.GetAsync(activity.From.AadObjectId, replyToId).ConfigureAwait(false);

                    if (!string.IsNullOrEmpty(replyToId))
                    {
                        var updateCardActivity = new Activity(ActivityTypes.Message)
                        {
                            Id = activityFromStorage.ActivityId,
                            Conversation = activity.Conversation,
                            Attachments = new List<Attachment> { attachment },
                        };
                        await turnContext.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
                    }

                    await turnContext.SendActivityAsync(MessageFactory.Text(string.Format(CultureInfo.CurrentCulture, Strings.RoomBooked, valuesFromTaskModule.RoomName)), cancellationToken).ConfigureAwait(false);
                }
                else
                {
                    if (!string.IsNullOrEmpty(replyToId))
                    {
                        await this.UpdateFavouriteCardAsync(turnContext, replyToId).ConfigureAwait(false);
                    }
                }

                return;
            }, default(CancellationToken));


            // Let the caller know proactive messages have been sent
            return this.Ok();
        }

        private async Task UpdateFavouriteCardAsync(ITurnContext turnContext, string activityReferenceId)
        {
            var activity = turnContext.Activity;
            var userAADToken = await this.tokenHelper.GetUserTokenAsync(activity.From.Id).ConfigureAwait(false);
            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in UpdateFavouriteCardAsync.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return;
            }

            var rooms = await this.favoriteStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            rooms = await this.meetingHelper.FilterFavoriteRoomsAsync(userAADToken, rooms?.ToList());
            var startUTCTime = DateTime.UtcNow.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var startTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
            var endTime = startTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);

            if (rooms?.Count > 0)
            {
                ScheduleRequest request = new ScheduleRequest
                {
                    StartDateTime = new DateTimeAndTimeZone() { DateTime = startTime, TimeZone = userConfiguration.IanaTimezone },
                    EndDateTime = new DateTimeAndTimeZone() { DateTime = endTime, TimeZone = userConfiguration.IanaTimezone },
                    Schedules = new List<string>(),
                };

                request.Schedules.AddRange(rooms.Select(room => room.RoomEmail));
                var roomsScheduleResponse = await this.meetingProvider.GetRoomsScheduleAsync(request, userAADToken).ConfigureAwait(false);
                if (roomsScheduleResponse.ErrorResponse == null)
                {
                    await this.SendAndUpdateCardAsync(turnContext, rooms, roomsScheduleResponse, activityReferenceId).ConfigureAwait(false);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(Strings.ExceptionResponse)).ConfigureAwait(false);
                }
            }
            else
            {
                RoomScheduleResponse scheduleResponse = new RoomScheduleResponse { Schedules = new List<Schedule>() };
                await this.SendAndUpdateCardAsync(turnContext, rooms, scheduleResponse, activityReferenceId).ConfigureAwait(false);
            }
        }

        private async Task SendAndUpdateCardAsync(ITurnContext turnContext, IList<UserFavoriteRoomEntity> rooms, RoomScheduleResponse scheduleResponse, string activityReferenceId)
        {
            var activity = turnContext.Activity;
            var startUTCTime = DateTime.UtcNow.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var endUTCTime = startUTCTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);
            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in SendAndUpdateCardAsync.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return;
            }

            var startTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));

            foreach (var room in scheduleResponse.Schedules)
            {
                var searchedRoom = rooms.Where(favoriteRoom => favoriteRoom.RowKey == room.ScheduleId).FirstOrDefault();
                room.RoomName = searchedRoom?.RoomName;
                room.BuildingName = searchedRoom?.BuildingName;
            }

            var activityFromStorage = await this.activityStorageProvider.GetAsync(turnContext.Activity.From.AadObjectId, activityReferenceId).ConfigureAwait(false);
            if (activityFromStorage != null)
            {
                var attachment = FavoriteRoomsListCard.GetFavoriteRoomsListAttachment(scheduleResponse, startUTCTime, endUTCTime, userConfiguration.WindowsTimezone, activityReferenceId);
                var updateCardActivity = new Activity(ActivityTypes.Message)
                {
                    Id = activityFromStorage.ActivityId,
                    Conversation = turnContext.Activity.Conversation,
                    Attachments = new List<Attachment> { attachment },
                };

                var activityResponse = await turnContext.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
                Models.TableEntities.ActivityEntity newActivity = new Models.TableEntities.ActivityEntity { ActivityId = activityResponse.Id, PartitionKey = turnContext.Activity.From.AadObjectId, RowKey = activityReferenceId };
                await this.activityStorageProvider.AddAsync(newActivity).ConfigureAwait(false);
            }
            else
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(Strings.FavoriteRoomsModified)).ConfigureAwait(false);
            }
        }

    }
}